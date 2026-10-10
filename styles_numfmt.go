package excelize

// GetStyleNumFmtCode provides a function to get the number format code of a
// style by the given style index, the same index the GetCellStyle function
// and the StyleID of a RowCell give. A custom number format gives its format
// code, and a built-in number format gives its format code under English
// localization, with the general number format given as "General". A number
// format that has no format code gives an empty string.
func (f *File) GetStyleNumFmtCode(idx int) (string, error) {
	f.mu.Lock()
	s, err := f.stylesReader()
	f.mu.Unlock()
	if err != nil {
		return "", err
	}
	if idx < 0 || s.CellXfs == nil || len(s.CellXfs.Xf) <= idx {
		return "", newInvalidStyleID(idx)
	}
	numFmtID := s.CellXfs.Xf[idx].NumFmtID
	if numFmtID == nil || *numFmtID == 0 {
		return "General", nil
	}
	if s.NumFmts != nil {
		for _, numFmt := range s.NumFmts.NumFmt {
			if numFmt != nil && numFmt.NumFmtID == *numFmtID {
				return numFmt.FormatCode, nil
			}
		}
	}
	return builtInNumFmt[*numFmtID], nil
}
