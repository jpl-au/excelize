package excelize

import (
	"fmt"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestGetSheetStructure(t *testing.T) {
	// Test get the worksheet structure against the functions that load the worksheet
	for _, opts := range []Options{{}, {UnzipXMLSizeLimit: 128}} {
		for _, name := range []string{"Book1.xlsx", "MergeCell.xlsx", "CalcChain.xlsx", "SharedStrings.xlsx"} {
			f, err := OpenFile(filepath.Join("test", name), opts)
			require.NoError(t, err)
			expected, err := OpenFile(filepath.Join("test", name))
			require.NoError(t, err)
			for _, sheet := range f.GetSheetList() {
				structure, err := f.GetSheetStructure(sheet)
				if _, ok := err.(ErrSheetNotExist); ok {
					continue
				}
				require.NoError(t, err, "%s %s", name, sheet)
				dimension, err := expected.GetSheetDimension(sheet)
				require.NoError(t, err)
				assert.Equal(t, dimension, structure.Dimension, "%s %s", name, sheet)
				mergeCells, err := expected.GetMergeCells(sheet, true)
				if err != nil {
					continue
				}
				var refs []string
				for _, mergeCell := range mergeCells {
					refs = append(refs, mergeCell[0])
				}
				for i, ref := range structure.MergeCells {
					// A merged range of one cell is stored as the cell
					if !strings.Contains(ref, ":") {
						structure.MergeCells[i] = ref + ":" + ref
					}
				}
				assert.Equal(t, refs, structure.MergeCells, "%s %s", name, sheet)
				tables, err := expected.GetTables(sheet)
				require.NoError(t, err)
				assert.Len(t, structure.TableParts, len(tables), "%s %s", name, sheet)
				for i, table := range tables {
					assert.Equal(t, table.tableXML, structure.TableParts[i])
				}
			}
			assert.NoError(t, expected.Close())
			assert.NoError(t, f.Close())
		}
	}
	// Test get the structure of a worksheet in memory
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "</sheetData>"))
	assert.NoError(t, f.MergeCell("Sheet1", "B2", "C3"))
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "A1", "https://github.com/xuri/excelize", "External", HyperlinkOpts{Display: new("Excelize")}))
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "A2", "Sheet1!B2", "Location"))
	assert.NoError(t, f.AddTable("Sheet1", &Table{Range: "E1:F3", Name: "Table1"}))
	assert.NoError(t, f.AddChart("Sheet1", "H1", &Chart{Type: Col, Series: []ChartSeries{{Name: "Sheet1!$E$1", Values: "Sheet1!$E$2:$E$3"}}}))
	style, err := f.NewStyle(&Style{NumFmt: 2})
	assert.NoError(t, err)
	assert.NoError(t, f.SetColStyle("Sheet1", "B", style))
	expected := SheetStructure{
		Cols:       []SheetCol{{Min: 2, Max: 2, StyleID: style}},
		MergeCells: []string{"B2:C3"},
		HyperLinks: []SheetHyperLink{
			{Ref: "A1", Link: "https://github.com/xuri/excelize", Display: "Excelize", External: true},
			{Ref: "A2", Link: "Sheet1!B2"},
		},
		TableParts:  []string{"xl/tables/table1.xml"},
		DrawingPart: "xl/drawings/drawing1.xml",
	}
	structure, err := f.GetSheetStructure("Sheet1")
	assert.NoError(t, err)
	expected.Dimension = structure.Dimension
	assert.Equal(t, expected, structure)
	// Test get the tables of a worksheet
	assert.NoError(t, f.AddTable("Sheet1", &Table{Range: "J1:K3", Name: "Table2", ShowHeaderRow: new(false)}))
	tables, err := f.GetSheetTables("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, []SheetTable{
		{Name: "Table1", Range: "E1:F3", Columns: []string{"Column1", "Column2"}, HeaderRow: true, AutoFilter: true},
		{Name: "Table2", Range: "J2:K3", Columns: []string{"Column1", "Column2"}},
	}, tables)
	expected.TableParts = append(expected.TableParts, "xl/tables/table2.xml")
	// Test get the tables of a worksheet with unsupported charset
	tableXML, _ := f.Pkg.Load("xl/tables/table2.xml")
	f.Pkg.Store("xl/tables/table2.xml", MacintoshCyrillicCharset)
	_, err = f.GetSheetTables("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	f.Pkg.Store("xl/tables/table2.xml", tableXML)
	// Test get the tables of a worksheet that does not exist
	_, err = f.GetSheetTables("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get the structure of a saved worksheet, in memory and in a temporary file
	path := filepath.Join("test", "TestGetSheetStructure.xlsx")
	assert.NoError(t, f.SaveAs(path))
	assert.NoError(t, f.Close())
	for _, opts := range []Options{{}, {UnzipXMLSizeLimit: 128}} {
		f, err = OpenFile(path, opts)
		assert.NoError(t, err)
		structure, err = f.GetSheetStructure("Sheet1")
		assert.NoError(t, err)
		assert.Equal(t, expected, structure)
		assert.NoError(t, f.Close())
	}
	// Test get the structure of a worksheet with empty rows, with a prefix on its elements and with rows that end in two chunks
	f = NewFile()
	for content, expected := range map[string]SheetStructure{
		`<worksheet xmlns="%s"><dimension ref="A1"/><sheetData/><mergeCells><mergeCell ref="A1:B1"/></mergeCells></worksheet>`:                                                          {Dimension: "A1", MergeCells: []string{"A1:B1"}},
		`<x:worksheet xmlns:x="%s"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData><x:mergeCells><x:mergeCell ref="A1:B1"/></x:mergeCells></x:worksheet>`: {MergeCells: []string{"A1:B1"}},
		`<worksheet xmlns="%s"><sheetData><row r="1"/></sheetData><mergeCells><mergeCell ref="A1:B1"/></mergeCells>` + strings.Repeat(" ", StreamChunkSize-40) + `</worksheet>`:         {MergeCells: []string{"A1:B1"}},
	} {
		f.Sheet.Delete("xl/worksheets/sheet1.xml")
		f.Pkg.Store("xl/worksheets/sheet1.xml", fmt.Appendf(nil, content, NameSpaceSpreadSheet.Value))
		structure, err = f.GetSheetStructure("Sheet1")
		assert.NoError(t, err)
		assert.Equal(t, expected, structure)
	}
	// Test get the structure of a worksheet with unsupported charset
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", MacintoshCyrillicCharset)
	_, err = f.GetSheetStructure("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	// Test get the structure of a worksheet with invalid name and of a worksheet that does not exist
	_, err = f.GetSheetStructure("Sheet:1")
	assert.Equal(t, ErrSheetNameInvalid, err)
	_, err = f.GetSheetStructure("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	assert.NoError(t, f.Close())
}
