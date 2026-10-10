package excelize

import "encoding/xml"

// RowCell is one cell of a worksheet row, as the Cells function of the rows
// iterator returns it.
type RowCell struct {
	// Col is the column number of the cell. The first column is 1.
	Col int
	// Type is the stored data type of the cell, the same GetCellType returns.
	Type CellType
	// Value is the value of the cell: formatted by default, or as stored when
	// the RawCellValue option is set. For a formula it is the cached result.
	Value string
	// Formula is the formula of the cell without the leading equal sign, or an
	// empty string when the cell has no formula. A cell that takes part in a
	// shared formula has the formula written out for that cell, the same text
	// GetCellFormula returns.
	Formula string
	// FormulaType is the stored type of the formula: STCellFormulaTypeArray,
	// STCellFormulaTypeDataTable, STCellFormulaTypeNormal or
	// STCellFormulaTypeShared, or an empty string when the cell has no formula
	// or the formula has no stored type.
	FormulaType string
	// StyleID is the style index of the cell, the same GetCellStyle returns
	// for a cell that has a style of its own, or 0.
	StyleID int
	// HasValue reports whether the cell stores a value, which may be an empty
	// string. For a formula it reports whether a cached result is stored.
	HasValue bool
}

// Cells returns the cells stored in the current row, in column order. A cell
// that stores only a style is one of them; it has an empty value and HasValue
// is false. It reads the worksheet as a stream in the same way as
// Columns, and gives the formula and the data type of each cell as well as
// its value, so that a large worksheet can be read in full without loading it
// into memory. Use Cells or Columns for a row, not both. This fetches the
// worksheet data as a stream, returns each cell in a row and will not skip
// empty rows in the tail of the worksheet. For example:
//
//	rows, err := f.Rows("Sheet1")
//	if err != nil {
//	    fmt.Println(err)
//	    return
//	}
//	for rows.Next() {
//	    cells, err := rows.Cells()
//	    if err != nil {
//	        fmt.Println(err)
//	    }
//	    for _, cell := range cells {
//	        fmt.Println(cell.Col, cell.Value, cell.Formula)
//	    }
//	}
//	if err = rows.Close(); err != nil {
//	    fmt.Println(err)
//	}
func (rows *Rows) Cells(opts ...Options) ([]RowCell, error) {
	if rows.curRow > rows.seekRow {
		return nil, nil
	}
	var (
		cells []RowCell
		col   int
		err   error
		token xml.Token
	)
	rows.rawCellValue = rows.f.getOptions(opts...).RawCellValue
	if rows.sst, err = rows.f.sharedStringsReader(); err != nil {
		return cells, err
	}
	for {
		if rows.token != nil {
			token = rows.token
		} else if token, _ = rows.decoder.Token(); token == nil {
			break
		}
		if xmlElement, ok := token.(xml.EndElement); ok && xmlElement.Name.Local == "sheetData" {
			return cells, err
		}
		xmlElement, ok := token.(xml.StartElement)
		if !ok {
			continue
		}
		if xmlElement.Name.Local == "row" {
			rowNum := 0
			if rowNum, err = attrValToInt("r", xmlElement.Attr); rowNum > TotalRows {
				rows.err, rows.token = ErrMaxRows, nil
				return cells, rows.err
			} else if rowNum != 0 {
				rows.curRow = rowNum
			} else if rows.token == nil {
				rows.curRow++
			}
			rows.token = token
			rows.seekRowOpts = extractRowOpts(xmlElement.Attr)
			if rows.curRow > rows.seekRow {
				rows.token = nil
				return cells, err
			}
		}
		if xmlElement.Name.Local == "c" {
			col++
			var cell RowCell
			if cell, err = rows.cell(&xmlElement, &col); err == nil {
				cells = append(cells, cell)
			}
		}
		rows.token = nil
		if err != nil {
			return cells, err
		}
	}
	return cells, err
}

// cell reads one cell element of the current row. It updates col from the cell
// reference when the element has one.
func (rows *Rows) cell(start *xml.StartElement, col *int) (RowCell, error) {
	c := xlsxC{XMLName: start.Name}
	err := c.cellXMLAttrHandler(start)
	if err != nil {
		return RowCell{}, err
	}
	hasValue := false
	for done := false; !done; {
		token, err := rows.decoder.Token()
		if err != nil {
			return RowCell{}, err
		}
		switch element := token.(type) {
		case xml.StartElement:
			switch element.Name.Local {
			case "v":
				hasValue = true
				err = rows.decoder.DecodeElement(&c.V, &element)
			case "f":
				err = rows.decoder.DecodeElement(&c.F, &element)
			case "is":
				hasValue = true
				err = rows.decoder.DecodeElement(&c.IS, &element)
			}
			if err != nil {
				return RowCell{}, err
			}
		case xml.EndElement:
			done = element == start.End()
		}
	}
	if c.R != "" {
		if *col, _, err = CellNameToCoordinates(c.R); err != nil {
			return RowCell{}, err
		}
	} else if c.R, err = CoordinatesToCellName(*col, rows.curRow); err != nil {
		return RowCell{}, err
	}
	cell := RowCell{Col: *col, Type: cellTypes[c.T], StyleID: c.S, HasValue: hasValue}
	if cell.Value, err = c.getValueFrom(rows.f, rows.sst, rows.rawCellValue); err != nil {
		return RowCell{}, err
	}
	if cell.Formula, err = rows.formula(&c); err != nil {
		return RowCell{}, err
	}
	if c.F != nil {
		cell.FormulaType = c.F.T
	}
	return cell, nil
}

// formula returns the formula of a cell read from the stream. A shared formula
// is stored once, in the first cell of its range; the cells that follow carry
// only its index, so the first cell is kept and the formula is converted for
// each cell that refers to it.
func (rows *Rows) formula(c *xlsxC) (string, error) {
	if c.F == nil {
		return "", nil
	}
	if c.F.T != STCellFormulaTypeShared || c.F.Si == nil {
		return c.F.Content, nil
	}
	if c.F.Ref != "" {
		if rows.sharedFormulas == nil {
			rows.sharedFormulas = make(map[int]*xlsxC)
		}
		rows.sharedFormulas[*c.F.Si] = c
	}
	first, ok := rows.sharedFormulas[*c.F.Si]
	if !ok {
		return "", nil
	}
	return first.convertSharedFormula(c.R)
}

// FormatCellValue returns the value of a cell with its number format applied,
// by the given cell that the Cells function read with the RawCellValue
// option. It is the value the Cells function returns for the cell without the
// option, so a worksheet can be read once and only the cells of interest
// formatted. The StyleID of the cell gives the number format, and can be
// changed first to apply the style of a row or of a column.
func (f *File) FormatCellValue(cell RowCell) (string, error) {
	c := xlsxC{S: cell.StyleID, V: cell.Value}
	switch cell.Type {
	case CellTypeBool:
		c.T = "b"
	case CellTypeDate:
		c.T = "d"
	case CellTypeError:
		c.T = "e"
	case CellTypeFormula:
		c.T = "str"
	case CellTypeInlineString, CellTypeSharedString:
		c.T, c.IS = "inlineStr", &xlsxSI{T: &xlsxT{Val: cell.Value}}
	}
	return c.getValueFrom(f, nil, false)
}
