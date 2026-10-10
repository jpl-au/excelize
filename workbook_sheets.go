package excelize

import (
	"archive/zip"
	"os"
	"path"
	"strings"
)

// WorkbookSheet is a sheet of the workbook, as the GetWorkbookSheets function
// returns it.
type WorkbookSheet struct {
	// Name is the name of the sheet.
	Name string
	// State is the visibility of the sheet: "visible", "hidden" or
	// "veryHidden".
	State string
	// Chartsheet reports whether the sheet is a chartsheet, which holds a
	// chart and no cells.
	Chartsheet bool
	// Part is the name of the part that holds the sheet, for example
	// "xl/worksheets/sheet1.xml".
	Part string
	// Size is the size of the part in bytes when it is not compressed.
	Size int64
}

// GetWorkbookSheets returns the worksheets and the chartsheets of the
// workbook in the order of the workbook, with the visibility, the part and
// the size of each. It does not read the sheets.
func (f *File) GetWorkbookSheets() ([]WorkbookSheet, error) {
	wb, err := f.workbookReader()
	if err != nil {
		return nil, err
	}
	rels, err := f.relsReader(f.getWorkbookRelsPath())
	if err != nil || rels == nil {
		return nil, err
	}
	rels.mu.Lock()
	defer rels.mu.Unlock()
	var sheets []WorkbookSheet
	for _, v := range wb.Sheets.Sheet {
		for _, rel := range rels.Relationships {
			if rel.ID != v.ID {
				continue
			}
			sheet := WorkbookSheet{
				Name:       v.Name,
				State:      v.State,
				Chartsheet: strings.HasSuffix(rel.Type, "/chartsheet"),
				Part:       f.getWorksheetPath(rel.Target),
			}
			if sheet.State == "" {
				sheet.State = "visible"
			}
			sheet.Size = f.partSize(sheet.Part)
			sheets = append(sheets, sheet)
		}
	}
	return sheets, nil
}

// partSize returns the size in bytes of a part of the workbook, without
// reading a part that was left in the ZIP archive on open.
func (f *File) partSize(name string) int64 {
	if part, ok := f.zipParts.Load(name); ok {
		return part.(*zip.File).FileInfo().Size()
	}
	if content, ok := f.Pkg.Load(name); ok && content != nil {
		return int64(len(content.([]byte)))
	}
	if tempFile, ok := f.tempFiles.Load(name); ok {
		if info, err := os.Stat(tempFile.(string)); err == nil {
			return info.Size()
		}
	}
	return 0
}

// GetExternalLinks returns the targets outside the workbook that the external
// link parts of the workbook refer to, for example the path of another
// workbook that a formula reads from.
func (f *File) GetExternalLinks() ([]string, error) {
	rels, err := f.relsReader(f.getWorkbookRelsPath())
	if err != nil || rels == nil {
		return nil, err
	}
	var parts []string
	rels.mu.Lock()
	for _, rel := range rels.Relationships {
		if strings.HasSuffix(rel.Type, "/externalLink") {
			parts = append(parts, f.getWorksheetPath(rel.Target))
		}
	}
	rels.mu.Unlock()
	var links []string
	for _, part := range parts {
		linkRels, err := f.relsReader(path.Dir(part) + "/_rels/" + path.Base(part) + ".rels")
		if err != nil {
			return links, err
		}
		if linkRels == nil {
			continue
		}
		linkRels.mu.Lock()
		for _, rel := range linkRels.Relationships {
			if rel.TargetMode == "External" {
				links = append(links, rel.Target)
			}
		}
		linkRels.mu.Unlock()
	}
	return links, nil
}
