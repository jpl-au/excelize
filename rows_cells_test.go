package excelize

import (
	"fmt"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// assertCellsMatch reads every worksheet of f with Cells and checks each cell
// against the functions that read one cell at a time, and each row against
// Columns. It returns the number of cells and of formulas read.
func assertCellsMatch(t *testing.T, f *File, opts ...Options) (cells, formulas int) {
	t.Helper()
	merged := func(sheet string) map[string]bool {
		inside := make(map[string]bool)
		ranges, err := f.GetMergeCells(sheet)
		require.NoError(t, err)
		for _, mergeCell := range ranges {
			rect, err := rangeRefToCoordinates(mergeCell[0])
			require.NoError(t, err)
			for col := rect[0]; col <= rect[2]; col++ {
				for row := rect[1]; row <= rect[3]; row++ {
					if col != rect[0] || row != rect[1] {
						name, _ := CoordinatesToCellName(col, row)
						inside[name] = true
					}
				}
			}
		}
		return inside
	}
	for _, sheet := range f.GetSheetList() {
		byCells, err := f.Rows(sheet)
		require.NoError(t, err)
		byColumns, err := f.Rows(sheet)
		require.NoError(t, err)
		inside := merged(sheet)
		for byCells.Next() {
			require.True(t, byColumns.Next())
			row, err := byCells.Cells(opts...)
			require.NoError(t, err)
			columns, err := byColumns.Columns(opts...)
			require.NoError(t, err)
			seen := make(map[int]bool)
			for _, cell := range row {
				name, err := CoordinatesToCellName(cell.Col, byCells.seekRow)
				require.NoError(t, err)
				require.False(t, seen[cell.Col], "%s!%s returned twice", sheet, name)
				seen[cell.Col] = true
				require.LessOrEqual(t, cell.Col, len(columns), "%s!%s", sheet, name)
				assert.Equal(t, columns[cell.Col-1], cell.Value, "%s!%s value against Columns", sheet, name)
				cells++
				if cell.Formula != "" {
					formulas++
				}
				// A cell inside a merged range is read by the single-cell
				// functions as the first cell of the range.
				if inside[name] {
					continue
				}
				value, err := f.GetCellValue(sheet, name, opts...)
				require.NoError(t, err)
				assert.Equal(t, value, cell.Value, "%s!%s value", sheet, name)
				formula, err := f.GetCellFormula(sheet, name)
				require.NoError(t, err)
				assert.Equal(t, formula, cell.Formula, "%s!%s formula", sheet, name)
				cellType, err := f.GetCellType(sheet, name)
				require.NoError(t, err)
				assert.Equal(t, cellType, cell.Type, "%s!%s type", sheet, name)
			}
			for index, value := range columns {
				if value != "" {
					assert.True(t, seen[index+1], "%s row %d column %d missing from Cells", sheet, byCells.seekRow, index+1)
				}
			}
		}
		require.False(t, byColumns.Next())
		require.NoError(t, byCells.Error())
		require.NoError(t, byCells.Close())
		require.NoError(t, byColumns.Close())
	}
	return cells, formulas
}

func TestFormatCellValue(t *testing.T) {
	// Test format the cells read as raw values against the cells read as formatted values
	for _, name := range []string{"Book1.xlsx", "CalcChain.xlsx", "SharedStrings.xlsx", "MergeCell.xlsx"} {
		f, err := OpenFile(filepath.Join("test", name))
		require.NoError(t, err)
		for _, sheet := range f.GetSheetList() {
			raw, err := f.Rows(sheet)
			if err != nil {
				continue
			}
			formatted, err := f.Rows(sheet)
			require.NoError(t, err)
			for raw.Next() {
				require.True(t, formatted.Next())
				cells, err := raw.Cells(Options{RawCellValue: true})
				require.NoError(t, err)
				expected, err := formatted.Cells()
				require.NoError(t, err)
				require.Len(t, cells, len(expected))
				for i, cell := range cells {
					value, err := f.FormatCellValue(cell)
					assert.NoError(t, err)
					assert.Equal(t, expected[i].Value, value, "%s %s row %d column %d", name, sheet, raw.seekRow, cell.Col)
				}
			}
			assert.NoError(t, raw.Close())
			assert.NoError(t, formatted.Close())
		}
		assert.NoError(t, f.Close())
	}
	// Test format a cell value with each data type and with the style of another cell
	f := NewFile()
	style, err := f.NewStyle(&Style{NumFmt: 2})
	assert.NoError(t, err)
	date, err := f.NewStyle(&Style{NumFmt: 14})
	assert.NoError(t, err)
	for expected, cell := range map[string]RowCell{
		"1.50":     {Type: CellTypeNumber, Value: "1.5", StyleID: style},
		"2.00":     {Value: "2", StyleID: style},
		"TRUE":     {Type: CellTypeBool, Value: "1"},
		"#N/A":     {Type: CellTypeError, Value: "#N/A"},
		"text":     {Type: CellTypeSharedString, Value: "text", StyleID: style},
		"inline":   {Type: CellTypeInlineString, Value: "inline"},
		"result":   {Type: CellTypeFormula, Value: "result"},
		"01-02-20": {Type: CellTypeDate, Value: "2020-01-02T00:00:00Z", StyleID: date},
		"":         {},
	} {
		value, err := f.FormatCellValue(cell)
		assert.NoError(t, err)
		assert.Equal(t, expected, value)
	}
	assert.NoError(t, f.Close())
}

func TestRowsCells(t *testing.T) {
	for _, name := range []string{"Book1.xlsx", "CalcChain.xlsx", "MergeCell.xlsx", "OverflowNumericCell.xlsx", "SharedStrings.xlsx"} {
		for _, raw := range []bool{false, true} {
			t.Run(fmt.Sprintf("%s raw %t", name, raw), func(t *testing.T) {
				f, err := OpenFile(filepath.Join("test", name))
				require.NoError(t, err)
				cells, _ := assertCellsMatch(t, f, Options{RawCellValue: raw})
				assert.NotZero(t, cells)
				assert.NoError(t, f.Close())
			})
		}
	}

	t.Run("shared formulas", func(t *testing.T) {
		f := NewFile()
		for row := 1; row <= 6; row++ {
			assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), row))
			assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), row*10))
		}
		formulaType, ref := STCellFormulaTypeShared, "C1:D6"
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "=A1+B1*$A$1+SUM($B1:B$6)", FormulaOpts{Ref: &ref, Type: &formulaType}))
		assert.NoError(t, f.SetCellFormula("Sheet1", "E2", "=A2&\"text\""))
		path := filepath.Join(t.TempDir(), "shared.xlsx")
		assert.NoError(t, f.SaveAs(path))
		assert.NoError(t, f.Close())

		f, err := OpenFile(path)
		require.NoError(t, err)
		_, formulas := assertCellsMatch(t, f, Options{RawCellValue: true})
		assert.Equal(t, 13, formulas)
		rows, err := f.Rows("Sheet1")
		require.NoError(t, err)
		var got []string
		for rows.Next() {
			cells, err := rows.Cells()
			require.NoError(t, err)
			for _, cell := range cells {
				if cell.Col == 4 {
					got = append(got, cell.Formula)
					assert.Equal(t, STCellFormulaTypeShared, cell.FormulaType)
				}
				if cell.Col < 3 || cell.Col == 5 {
					assert.Empty(t, cell.FormulaType)
				}
			}
		}
		assert.Equal(t, "B1+C1*$A$1+SUM($B1:C$6)", got[0])
		assert.Equal(t, "B6+C6*$A$1+SUM($B6:C$6)", got[5])
		assert.NoError(t, rows.Close())
		assert.NoError(t, f.Close())
	})

	t.Run("cells without references", func(t *testing.T) {
		f := NewFile()
		f.Sheet.Delete("xl/worksheets/sheet1.xml")
		f.Pkg.Store("xl/worksheets/sheet1.xml", fmt.Appendf(nil, `<worksheet xmlns="%s"><sheetData><row><c><v>1</v></c><c t="str"><f>A1*2</f><v>2</v></c></row><row r="4"><c r="C4" t="inlineStr"><is><t>text</t></is></c><c s="1"/></row></sheetData></worksheet>`, NameSpaceSpreadSheet.Value))
		rows, err := f.Rows("Sheet1")
		require.NoError(t, err)
		var got [][]RowCell
		for rows.Next() {
			cells, err := rows.Cells()
			require.NoError(t, err)
			got = append(got, cells)
		}
		assert.Equal(t, [][]RowCell{
			{{Col: 1, Value: "1", HasValue: true}, {Col: 2, Type: CellTypeFormula, Value: "2", Formula: "A1*2", HasValue: true}},
			nil, nil,
			{{Col: 3, Type: CellTypeInlineString, Value: "text", HasValue: true}, {Col: 4, StyleID: 1}},
		}, got)
		assert.NoError(t, rows.Close())
		assert.NoError(t, f.Close())
	})

	t.Run("after the last row", func(t *testing.T) {
		f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
		require.NoError(t, err)
		rows, err := f.Rows("Sheet2")
		require.NoError(t, err)
		cells, err := rows.Cells()
		assert.NoError(t, err)
		assert.Nil(t, cells)
		assert.NoError(t, rows.Close())
		assert.NoError(t, f.Close())
	})

	t.Run("invalid cell reference", func(t *testing.T) {
		f := NewFile()
		f.Sheet.Delete("xl/worksheets/sheet1.xml")
		f.Pkg.Store("xl/worksheets/sheet1.xml", fmt.Appendf(nil, `<worksheet xmlns="%s"><sheetData><row r="1"><c r="A"><v>1</v></c></row></sheetData></worksheet>`, NameSpaceSpreadSheet.Value))
		rows, err := f.Rows("Sheet1")
		require.NoError(t, err)
		assert.True(t, rows.Next())
		_, err = rows.Cells()
		assert.Error(t, err)
		assert.NoError(t, rows.Close())
		assert.NoError(t, f.Close())
	})
}

func BenchmarkRowsCells(b *testing.B) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	require.NoError(b, err)
	for b.Loop() {
		rows, err := f.Rows("Sheet2")
		require.NoError(b, err)
		for rows.Next() {
			if _, err := rows.Cells(); err != nil {
				b.Fatal(err)
			}
		}
		require.NoError(b, rows.Close())
	}
	require.NoError(b, f.Close())
}
