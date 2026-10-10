package excelize

import (
	"bytes"
	"encoding/xml"
	"io"
	"path"
	"runtime"
	"slices"
	"strconv"
	"strings"
	"sync"
)

// SheetChart is a chart of a worksheet or of a chartsheet, as the
// GetSheetCharts function returns it.
type SheetChart struct {
	// Name is the name of the chart in its drawing, or the name of its part
	// when the drawing gives it none, for example "chart1".
	Name string
	// Type is the type of the chart as the part names it, for example "bar"
	// or "scatter". A chart that combines several types gives them separated
	// by commas, and a part that holds no chart type gives an empty string.
	Type string
	// Title is the text of the title of the chart.
	Title string
	// Cell is the cell that the top left corner of the chart is anchored to,
	// or an empty string when the chart is not anchored to a cell.
	Cell string
	// Series are the data series of the chart.
	Series []SheetChartSeries
}

// SheetChartSeries is a data series of a chart and the formulas that give its
// data.
type SheetChartSeries struct {
	// Name is the stored text of the name of the series.
	Name string
	// NameRef is the formula of the name of the series.
	NameRef string
	// CategoryRef is the formula of the categories or of the X values.
	CategoryRef string
	// ValueRef is the formula of the values or of the Y values.
	ValueRef string
	// BubbleSizeRef is the formula of the bubble sizes.
	BubbleSizeRef string
}

// sheetChartAnchor is a chart that a drawing anchors.
type sheetChartAnchor struct {
	rID, name, cell string
}

// GetSheetCharts returns the charts of a worksheet or of a chartsheet, in
// the order of the drawing, with the type, the title, the anchor cell and the
// series formulas of each. It does not load the worksheet into memory.
func (f *File) GetSheetCharts(sheet string) ([]SheetChart, error) {
	structure, err := f.GetSheetStructure(sheet)
	if err != nil || structure.DrawingPart == "" {
		return nil, err
	}
	anchors, err := f.sheetChartAnchors(structure.DrawingPart)
	if err != nil {
		return nil, err
	}
	rels, _ := f.relsReader(path.Dir(structure.DrawingPart) + "/_rels/" + path.Base(structure.DrawingPart) + ".rels")
	var parts []string
	for i := 0; i < len(anchors); i++ {
		part := ""
		if rels != nil {
			rels.mu.Lock()
			for _, rel := range rels.Relationships {
				if rel.ID == anchors[i].rID && strings.HasSuffix(rel.Type, "/chart") {
					part = strings.TrimPrefix(path.Clean(rel.Target), "/")
					if !strings.HasPrefix(rel.Target, "/") {
						part = path.Join(path.Dir(structure.DrawingPart), rel.Target)
					}
				}
			}
			rels.mu.Unlock()
		}
		if part == "" {
			anchors = slices.Delete(anchors, i, i+1)
			i--
			continue
		}
		parts = append(parts, part)
	}
	var (
		charts = make([]SheetChart, len(anchors))
		errs   = make([]error, len(anchors))
		wg     sync.WaitGroup
		slots  = make(chan struct{}, runtime.GOMAXPROCS(0))
	)
	for i, part := range parts {
		wg.Go(func() {
			slots <- struct{}{}
			defer func() { <-slots }()
			charts[i], errs[i] = f.sheetChart(part)
			charts[i].Cell = anchors[i].cell
			if charts[i].Name = anchors[i].name; charts[i].Name == "" {
				charts[i].Name = strings.TrimSuffix(path.Base(part), ".xml")
			}
		})
	}
	wg.Wait()
	for _, err := range errs {
		if err != nil {
			return nil, err
		}
	}
	return charts, nil
}

// sheetChartAnchors reads the charts that a drawing part anchors.
func (f *File) sheetChartAnchors(part string) ([]sheetChartAnchor, error) {
	var (
		anchors  []sheetChartAnchor
		current  *sheetChartAnchor
		stack    []string
		row, col = -1, -1
		decoder  = f.xmlNewDecoder(bytes.NewReader(f.readXML(part)))
	)
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return anchors, nil
		}
		if err != nil {
			return nil, err
		}
		switch element := token.(type) {
		case xml.StartElement:
			stack = append(stack, element.Name.Local)
			switch element.Name.Local {
			case "oneCellAnchor", "twoCellAnchor", "absoluteAnchor":
				current, row, col = &sheetChartAnchor{}, -1, -1
			case "chart":
				if current != nil {
					current.rID = attrVal("id", element.Attr)
				}
			case "cNvPr":
				if current != nil {
					current.name = attrVal("name", element.Attr)
				}
			}
		case xml.CharData:
			if current == nil || len(stack) < 2 || stack[len(stack)-2] != "from" {
				continue
			}
			switch stack[len(stack)-1] {
			case "row":
				if row, err = strconv.Atoi(strings.TrimSpace(string(element))); err != nil {
					return nil, err
				}
			case "col":
				if col, err = strconv.Atoi(strings.TrimSpace(string(element))); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			switch element.Name.Local {
			case "oneCellAnchor", "twoCellAnchor", "absoluteAnchor":
				if current != nil && current.rID != "" {
					if row >= 0 && col >= 0 {
						if current.cell, err = CoordinatesToCellName(col+1, row+1); err != nil {
							return nil, err
						}
					}
					anchors = append(anchors, *current)
				}
				current = nil
			}
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		}
	}
}

// sheetChart reads the type, the title and the series of a chart part.
func (f *File) sheetChart(part string) (SheetChart, error) {
	var (
		chart   SheetChart
		types   []string
		stack   []string
		series  *SheetChartSeries
		decoder = f.xmlNewDecoder(bytes.NewReader(f.readXML(part)))
	)
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			chart.Type = strings.Join(types, ",")
			return chart, nil
		}
		if err != nil {
			return chart, err
		}
		switch element := token.(type) {
		case xml.StartElement:
			name := element.Name.Local
			stack = append(stack, name)
			if chartType, ok := strings.CutSuffix(name, "Chart"); ok && name != "chart" && !slices.Contains(types, chartType) {
				types = append(types, chartType)
			}
			if name == "ser" {
				series = &SheetChartSeries{}
			}
		case xml.CharData:
			value := strings.TrimSpace(string(element))
			if len(stack) == 0 || value == "" {
				continue
			}
			last := stack[len(stack)-1]
			if last == "t" {
				// The title text is inside the title of the chart, not of an axis
				for i := 1; i < len(stack); i++ {
					if stack[i-1] == "chart" && stack[i] == "title" {
						chart.Title += string(element)
						break
					}
				}
			}
			if series == nil {
				continue
			}
			if last == "v" && slices.Contains(stack, "tx") {
				series.Name += value
			}
			if last == "f" {
				switch {
				case slices.Contains(stack, "tx"):
					series.NameRef += value
				case slices.Contains(stack, "cat"), slices.Contains(stack, "xVal"):
					series.CategoryRef += value
				case slices.Contains(stack, "val"), slices.Contains(stack, "yVal"):
					series.ValueRef += value
				case slices.Contains(stack, "bubbleSize"):
					series.BubbleSizeRef += value
				}
			}
		case xml.EndElement:
			if element.Name.Local == "ser" && series != nil {
				chart.Series = append(chart.Series, *series)
				series = nil
			}
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		}
	}
}
