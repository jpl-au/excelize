package excelize

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestPartChecksums(t *testing.T) {
	dir := t.TempDir()
	save := func(name string, edit func(f *File)) string {
		f := NewFile()
		_, err := f.NewSheet("Sheet2")
		require.NoError(t, err)
		require.NoError(t, f.SetCellValue("Sheet1", "A1", 1))
		require.NoError(t, f.SetCellValue("Sheet2", "A1", 2))
		edit(f)
		path := filepath.Join(dir, name)
		require.NoError(t, f.SaveAs(path))
		require.NoError(t, f.Close())
		return path
	}
	before := save("before.xlsx", func(*File) {})
	same := save("same.xlsx", func(*File) {})
	after := save("after.xlsx", func(f *File) { require.NoError(t, f.SetCellValue("Sheet2", "A1", 3)) })

	beforeSums, err := PartChecksums(before)
	require.NoError(t, err)
	sameSums, err := PartChecksums(same)
	require.NoError(t, err)
	afterSums, err := PartChecksums(after)
	require.NoError(t, err)

	f, err := OpenFile(after)
	require.NoError(t, err)
	sheet1, err := f.SheetPart("Sheet1")
	require.NoError(t, err)
	sheet2, err := f.SheetPart("sheet2")
	require.NoError(t, err)
	assert.Equal(t, "xl/worksheets/sheet1.xml", sheet1)
	assert.Contains(t, beforeSums, sheet1)
	assert.Contains(t, beforeSums, sheet2)
	assert.Contains(t, beforeSums, "[Content_Types].xml")
	assert.Equal(t, beforeSums[sheet1], sameSums[sheet1])
	assert.Equal(t, beforeSums[sheet2], sameSums[sheet2])
	assert.Equal(t, beforeSums[sheet1], afterSums[sheet1])
	assert.NotEqual(t, beforeSums[sheet2], afterSums[sheet2])

	// Test get the part of a worksheet which does not exist, and with an
	// invalid sheet name
	_, err = f.SheetPart("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	_, err = f.SheetPart("Sheet:1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
	assert.NoError(t, f.Close())

	// Test get the checksums of a file which does not exist, of a file which
	// is not a spreadsheet, and of an encrypted spreadsheet
	_, err = PartChecksums(filepath.Join(dir, "missing.xlsx"))
	assert.ErrorIs(t, err, os.ErrNotExist)
	text := filepath.Join(dir, "text.xlsx")
	require.NoError(t, os.WriteFile(text, []byte("text"), 0o600))
	_, err = PartChecksums(text)
	assert.Error(t, err)
	_, err = PartChecksums(filepath.Join("test", "encryptAES.xlsx"))
	assert.Error(t, err)

	assert.Equal(t, "xl/sharedStrings.xml", partName(`xl\SharedStrings.xml`))
	assert.Equal(t, "[Content_Types].xml", partName("[content_types].xml"))
}
