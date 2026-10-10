package excelize

import (
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestGetWorkbookSheets(t *testing.T) {
	f := NewFile()
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	_, err = f.NewSheet("Sheet3")
	assert.NoError(t, err)
	assert.NoError(t, f.SetSheetVisible("Sheet2", false))
	assert.NoError(t, f.SetSheetVisible("Sheet3", false, true))
	assert.NoError(t, f.AddChartSheet("Chart1", &Chart{Type: Col, Series: []ChartSeries{{Values: "Sheet1!$A$1:$A$2"}}}))
	path := filepath.Join("test", "TestGetWorkbookSheets.xlsx")
	assert.NoError(t, f.SaveAs(path))
	assert.NoError(t, f.Close())
	// Test get the sheets of a workbook, before and after the sheets are read
	for _, opts := range []Options{{}, {UnzipXMLSizeLimit: 128}} {
		f, err = OpenFile(path, opts)
		assert.NoError(t, err)
		for range 2 {
			sheets, err := f.GetWorkbookSheets()
			assert.NoError(t, err)
			assert.Len(t, sheets, 4)
			for i, sheet := range sheets {
				assert.Positive(t, sheet.Size)
				sheets[i].Size = 0
			}
			assert.Equal(t, []WorkbookSheet{
				{Name: "Sheet1", State: "visible", Part: "xl/worksheets/sheet1.xml"},
				{Name: "Sheet2", State: "hidden", Part: "xl/worksheets/sheet2.xml"},
				{Name: "Sheet3", State: "veryHidden", Part: "xl/worksheets/sheet3.xml"},
				{Name: "Chart1", State: "visible", Chartsheet: true, Part: "xl/chartsheets/sheet4.xml"},
			}, sheets)
			_, err = f.GetRows("Sheet1")
			assert.NoError(t, err)
		}
		assert.NoError(t, f.Close())
	}
	// Test get the sheets and the external links of a workbook without relationships
	f = NewFile()
	f.Relationships.Delete(defaultXMLPathWorkbookRels)
	f.Pkg.Delete(defaultXMLPathWorkbookRels)
	sheets, err := f.GetWorkbookSheets()
	assert.NoError(t, err)
	assert.Empty(t, sheets)
	links, err := f.GetExternalLinks()
	assert.NoError(t, err)
	assert.Empty(t, links)
	assert.Zero(t, f.partSize("xl/worksheets/sheetN.xml"))
	// Test get the sheets and the external links of a workbook with unsupported charset
	f.Pkg.Store(defaultXMLPathWorkbookRels, MacintoshCyrillicCharset)
	_, err = f.GetWorkbookSheets()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	_, err = f.GetExternalLinks()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	_, err = f.GetWorkbookSheets()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}

func TestGetExternalLinks(t *testing.T) {
	f := NewFile()
	// Test get the external links of a workbook that has none
	links, err := f.GetExternalLinks()
	assert.NoError(t, err)
	assert.Empty(t, links)
	// Test get the external links of a workbook
	rels, err := f.relsReader(defaultXMLPathWorkbookRels)
	assert.NoError(t, err)
	rels.Relationships = append(rels.Relationships,
		xlsxRelationship{ID: "rId90", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink", Target: "externalLinks/externalLink1.xml"},
		xlsxRelationship{ID: "rId91", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink", Target: "externalLinks/externalLink2.xml"},
	)
	f.Pkg.Store("xl/externalLinks/_rels/externalLink1.xml.rels", []byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="Book2.xlsx" TargetMode="External"/></Relationships>`))
	links, err = f.GetExternalLinks()
	assert.NoError(t, err)
	assert.Equal(t, []string{"Book2.xlsx"}, links)
	// Test get the external links with unsupported charset relationships
	f.Relationships.Delete("xl/externalLinks/_rels/externalLink1.xml.rels")
	f.Pkg.Store("xl/externalLinks/_rels/externalLink1.xml.rels", MacintoshCyrillicCharset)
	_, err = f.GetExternalLinks()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}
