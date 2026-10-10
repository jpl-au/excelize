package excelize

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestGetSheetCharts(t *testing.T) {
	f := NewFile()
	for cell, value := range map[string]any{"A1": "Name", "B1": "Amount", "A2": "East", "B2": 2, "A3": "West", "B3": 3} {
		assert.NoError(t, f.SetCellValue("Sheet1", cell, value))
	}
	series := []ChartSeries{{Name: "Sheet1!$B$1", Categories: "Sheet1!$A$2:$A$3", Values: "Sheet1!$B$2:$B$3"}}
	assert.NoError(t, f.AddChart("Sheet1", "D2", &Chart{Type: Col, Series: series, Title: ChartTitle{Paragraph: []RichTextRun{{Text: "Amount by name"}}}}))
	assert.NoError(t, f.AddChart("Sheet1", "D20", &Chart{Type: Line, Series: series}, &Chart{Type: Scatter, Series: series}))
	assert.NoError(t, f.AddChartSheet("Chart1", &Chart{Type: Bubble, Series: []ChartSeries{{Name: "Sheet1!$B$1", Categories: "Sheet1!$A$2:$A$3", Values: "Sheet1!$B$2:$B$3", Sizes: "Sheet1!$B$2:$B$3"}}}))
	path := filepath.Join("test", "TestGetSheetCharts.xlsx")
	assert.NoError(t, f.SaveAs(path))
	assert.NoError(t, f.Close())
	f, err := OpenFile(path)
	assert.NoError(t, err)
	// Test get the charts of a worksheet
	charts, err := f.GetSheetCharts("Sheet1")
	assert.NoError(t, err)
	expected := []SheetChartSeries{{NameRef: "Sheet1!$B$1", CategoryRef: "Sheet1!$A$2:$A$3", ValueRef: "Sheet1!$B$2:$B$3"}}
	assert.Equal(t, []SheetChart{
		{Name: "Chart 2", Type: "bar", Title: "Amount by name", Cell: "D2", Series: expected},
		{Name: "Chart 3", Type: "line,scatter", Cell: "D20", Series: append(expected, expected...)},
	}, charts)
	// Test get the charts of a chartsheet
	charts, err = f.GetSheetCharts("Chart1")
	assert.NoError(t, err)
	assert.Len(t, charts, 1)
	assert.Equal(t, "bubble", charts[0].Type)
	assert.Equal(t, "Sheet1!$B$2:$B$3", charts[0].Series[0].BubbleSizeRef)
	// Test get the charts of a worksheet without a relationship for a chart
	f.Relationships.Delete("xl/drawings/_rels/drawing1.xml.rels")
	f.Pkg.Delete("xl/drawings/_rels/drawing1.xml.rels")
	charts, err = f.GetSheetCharts("Sheet1")
	assert.NoError(t, err)
	assert.Empty(t, charts)
	// Test get the charts of a chartsheet with unsupported charset chart and drawing
	f.Pkg.Range(func(k, _ any) bool {
		if strings.HasPrefix(k.(string), "xl/charts/chart") {
			f.Pkg.Store(k, MacintoshCyrillicCharset)
		}
		return true
	})
	_, err = f.GetSheetCharts("Chart1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	f.Pkg.Store("xl/drawings/drawing2.xml", MacintoshCyrillicCharset)
	_, err = f.GetSheetCharts("Chart1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	// Test get the charts of a worksheet without a drawing and of a worksheet that does not exist
	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	charts, err = f.GetSheetCharts("Sheet2")
	assert.NoError(t, err)
	assert.Empty(t, charts)
	_, err = f.GetSheetCharts("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	assert.NoError(t, f.Close())
}
