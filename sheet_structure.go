package excelize

import (
	"bytes"
	"encoding/xml"
	"io"
	"os"
	"path"
	"strings"
)

// SheetStructure is what a worksheet stores besides the cells of its rows, as
// the GetSheetStructure function returns it.
type SheetStructure struct {
	// Dimension is the stored used range of the worksheet, the same
	// GetSheetDimension returns.
	Dimension string
	// Cols are the column ranges of the worksheet that have a style.
	Cols []SheetCol
	// MergeCells are the references of the merged cell ranges, for example
	// "A1:B2".
	MergeCells []string
	// HyperLinks are the hyperlinks of the worksheet.
	HyperLinks []SheetHyperLink
	// TableParts are the names of the parts that hold the tables of the
	// worksheet, for example "xl/tables/table1.xml".
	TableParts []string
	// DrawingPart is the name of the part that holds the drawing of the
	// worksheet, for example "xl/drawings/drawing1.xml", or an empty string.
	DrawingPart string
}

// SheetCol is a range of columns of a worksheet and the style index the
// columns have. The first column is 1.
type SheetCol struct {
	Min, Max, StyleID int
}

// SheetHyperLink is a hyperlink of a worksheet.
type SheetHyperLink struct {
	// Ref is the cell or the range of cells that holds the hyperlink.
	Ref string
	// Link is the target of the hyperlink: an address outside the workbook
	// when External is true, otherwise a location in the workbook.
	Link string
	// Display is the stored display text of the hyperlink.
	Display string
	// External reports whether the hyperlink leads outside the workbook.
	External bool
}

// sheetDataEnd is the end of the rows of a worksheet. The character data of a
// cell cannot hold it, because the less than sign is escaped there.
var sheetDataEnd = []byte("</sheetData>")

// GetSheetStructure returns the stored dimension, the column styles, the
// merged cell ranges, the hyperlinks and the table and drawing parts of a
// worksheet. It does not load the worksheet into memory and does not parse
// the rows, so it is fast on a worksheet of any size. The rows of the
// worksheet can be read with the Rows function.
func (f *File) GetSheetStructure(sheet string) (SheetStructure, error) {
	var structure SheetStructure
	if err := checkSheetName(sheet); err != nil {
		return structure, err
	}
	name, ok := f.getSheetXMLPath(sheet)
	if !ok {
		return structure, ErrSheetNotExist{sheet}
	}
	if worksheet, ok := f.Sheet.Load(name); ok && worksheet != nil {
		ws := worksheet.(*xlsxWorksheet)
		ws.mu.Lock()
		// Flush data
		output, _ := xml.Marshal(ws)
		f.saveFileList(name, f.replaceNameSpaceBytes(name, output))
		ws.mu.Unlock()
	}
	var (
		rIDs     sheetStructureRIDs
		source   io.ReaderAt
		size     int64
		tempFile *os.File
		err      error
	)
	if content := f.readXML(name); len(content) > 0 {
		source, size = bytes.NewReader(content), int64(len(content))
	} else {
		if tempFile, err = f.readTemp(name); err != nil || tempFile == nil {
			return structure, err
		}
		defer func() {
			_ = tempFile.Close()
		}()
		info, err := tempFile.Stat()
		if err != nil {
			return structure, err
		}
		source, size = tempFile, info.Size()
	}
	decoder := f.xmlNewDecoder(io.NewSectionReader(source, 0, size))
	rows, err := structure.read(decoder, &rIDs, true)
	if err != nil {
		return structure, err
	}
	if rows {
		// Skip the rows without parsing them
		tail, err := sheetDataTail(source, decoder.InputOffset(), size)
		if err != nil {
			return structure, err
		}
		rest, restRIDs := structure, rIDs
		if _, err = rest.read(f.xmlNewDecoder(io.MultiReader(strings.NewReader("<worksheet>"), tail)), &restRIDs, false); err == nil {
			structure, rIDs = rest, restRIDs
		} else {
			// Parse the rows when the worksheet cannot be read in two parts
			if err = decoder.Skip(); err != nil {
				return structure, err
			}
			if _, err = structure.read(decoder, &rIDs, false); err != nil {
				return structure, err
			}
		}
	}
	rels, _ := f.relsReader(path.Dir(name) + "/_rels/" + path.Base(name) + ".rels")
	target := func(rID string) (string, bool, bool) {
		if rels == nil {
			return "", false, false
		}
		rels.mu.Lock()
		defer rels.mu.Unlock()
		for _, rel := range rels.Relationships {
			if rel.ID == rID {
				return rel.Target, rel.TargetMode == "External", true
			}
		}
		return "", false, false
	}
	part := func(rID string) string {
		target, external, ok := target(rID)
		if !ok || external {
			return ""
		}
		if strings.HasPrefix(target, "/") {
			return strings.TrimPrefix(path.Clean(target), "/")
		}
		return path.Join(path.Dir(name), target)
	}
	for i, rID := range rIDs.hyperLinks {
		if rID != "" {
			structure.HyperLinks[i].Link, structure.HyperLinks[i].External, _ = target(rID)
		}
	}
	for _, rID := range rIDs.tableParts {
		if tablePart := part(rID); tablePart != "" {
			structure.TableParts = append(structure.TableParts, tablePart)
		}
	}
	structure.DrawingPart = part(rIDs.drawing)
	return structure, nil
}

// sheetStructureRIDs holds the relationship IDs read from a worksheet. The
// hyperLinks have one entry for each hyperlink, empty for a location in the
// workbook.
type sheetStructureRIDs struct {
	hyperLinks []string
	tableParts []string
	drawing    string
}

// read reads the elements of a worksheet from the decoder. When stopAtRows is
// true it returns true at the start of the rows.
func (structure *SheetStructure) read(decoder *xml.Decoder, rIDs *sheetStructureRIDs, stopAtRows bool) (bool, error) {
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return false, nil
		}
		if err != nil {
			return false, err
		}
		element, ok := token.(xml.StartElement)
		if !ok {
			continue
		}
		switch element.Name.Local {
		case "sheetData":
			if stopAtRows {
				return true, nil
			}
		case "dimension":
			structure.Dimension = attrVal("ref", element.Attr)
		case "col":
			col := SheetCol{}
			col.Min, _ = attrValToInt("min", element.Attr)
			col.Max, _ = attrValToInt("max", element.Attr)
			if col.StyleID, _ = attrValToInt("style", element.Attr); col.StyleID != 0 {
				structure.Cols = append(structure.Cols, col)
			}
		case "mergeCell":
			structure.MergeCells = append(structure.MergeCells, attrVal("ref", element.Attr))
		case "hyperlink":
			structure.HyperLinks = append(structure.HyperLinks, SheetHyperLink{
				Ref:     attrVal("ref", element.Attr),
				Link:    attrVal("location", element.Attr),
				Display: attrVal("display", element.Attr),
			})
			rIDs.hyperLinks = append(rIDs.hyperLinks, attrVal("id", element.Attr))
		case "tablePart":
			rIDs.tableParts = append(rIDs.tableParts, attrVal("id", element.Attr))
		case "drawing":
			rIDs.drawing = attrVal("id", element.Attr)
		}
	}
}

// attrVal returns the value of the attribute with the given local name.
func attrVal(name string, attrs []xml.Attr) string {
	for _, attr := range attrs {
		if attr.Name.Local == name {
			return attr.Value
		}
	}
	return ""
}

// sheetDataTail returns a reader of the worksheet that starts after the end
// of the rows, searching backwards from the end of the worksheet. A worksheet
// with no end of rows after the given offset has empty rows, and the reader
// starts at the offset.
func sheetDataTail(source io.ReaderAt, offset, size int64) (io.Reader, error) {
	buf := make([]byte, min(size-offset, StreamChunkSize))
	for end := size; end > offset; {
		start := max(end-int64(len(buf)), offset)
		chunk := buf[:end-start]
		if _, err := source.ReadAt(chunk, start); err != nil && err != io.EOF {
			return nil, err
		}
		if i := bytes.LastIndex(chunk, sheetDataEnd); i >= 0 {
			from := start + int64(i+len(sheetDataEnd))
			return io.NewSectionReader(source, from, size-from), nil
		}
		if start == offset {
			break
		}
		// Keep enough of the chunk to find an end that spans two chunks
		end = start + int64(len(sheetDataEnd)) - 1
	}
	return io.NewSectionReader(source, offset, size-offset), nil
}

// SheetTable is a table of a worksheet, as the GetSheetTables function returns
// it.
type SheetTable struct {
	// Name is the display name of the table, or its name when it has no
	// display name.
	Name string
	// Range is the range of cells the table covers, with its header row and
	// its totals rows.
	Range string
	// Columns are the names of the columns of the table.
	Columns []string
	// HeaderRow reports whether the table has a header row.
	HeaderRow bool
	// TotalsRowCount is the number of totals rows at the end of the table.
	TotalsRowCount int
	// AutoFilter reports whether the table has an auto filter.
	AutoFilter bool
}

// GetSheetTables returns the tables of a worksheet with their columns. It
// does not load the worksheet into memory, so it is fast on a worksheet of
// any size.
func (f *File) GetSheetTables(sheet string) ([]SheetTable, error) {
	structure, err := f.GetSheetStructure(sheet)
	if err != nil {
		return nil, err
	}
	var tables []SheetTable
	for _, part := range structure.TableParts {
		content := f.readXML(part)
		if len(content) == 0 {
			continue
		}
		var t xlsxTable
		if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(content))).
			Decode(&t); err != nil && err != io.EOF {
			return tables, err
		}
		table := SheetTable{
			Name:           t.DisplayName,
			Range:          t.Ref,
			HeaderRow:      t.HeaderRowCount == nil || *t.HeaderRowCount > 0,
			TotalsRowCount: t.TotalsRowCount,
			AutoFilter:     t.AutoFilter != nil,
		}
		if table.Name == "" {
			table.Name = t.Name
		}
		if t.TableColumns != nil {
			for _, column := range t.TableColumns.TableColumn {
				if column != nil {
					table.Columns = append(table.Columns, column.Name)
				}
			}
		}
		tables = append(tables, table)
	}
	return tables, nil
}
