package excelize

import (
	"archive/zip"
	"strings"
)

// PartChecksums returns the CRC-32 checksum of each part stored in the
// spreadsheet file with the given name, by part name. The checksums are read
// from the directory of the file without reading the parts, so two versions
// of a large workbook can be compared part by part before either is opened: a
// part with the same checksum in both files has the same content. The part of
// a worksheet is given by SheetPart. An encrypted spreadsheet is not stored as
// parts and returns an error.
func PartChecksums(name string) (map[string]uint32, error) {
	archive, err := zip.OpenReader(name)
	if err != nil {
		return nil, err
	}
	checksums := make(map[string]uint32, len(archive.File))
	for _, file := range archive.File {
		checksums[partName(file.Name)] = file.CRC32
	}
	return checksums, archive.Close()
}

// partName returns the name ReadZipReader gives to the part stored under the
// given name in the spreadsheet file.
func partName(name string) string {
	name = strings.ReplaceAll(name, "\\", "/")
	switch strings.ToLower(name) {
	case strings.ToLower(defaultXMLPathContentTypes):
		return defaultXMLPathContentTypes
	case strings.ToLower(defaultXMLPathSharedStrings):
		return defaultXMLPathSharedStrings
	}
	return name
}

// SheetPart returns the name of the part that stores the worksheet with the
// given name, for example xl/worksheets/sheet1.xml.
func (f *File) SheetPart(sheet string) (string, error) {
	if err := checkSheetName(sheet); err != nil {
		return "", err
	}
	name, ok := f.getSheetXMLPath(sheet)
	if !ok {
		return "", ErrSheetNotExist{sheet}
	}
	return name, nil
}

// SharedStringsPart is the name of the part that stores the text shared by the
// cells of every worksheet. A worksheet part with the same checksum in two
// files has the same cells only when this part has the same checksum too.
const SharedStringsPart = defaultXMLPathSharedStrings
