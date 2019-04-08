package hoboexcel

import (
	"archive/zip"
	"bufio"
	"encoding/xml"
	"fmt"
	"html"
	"io"
	"os"
	"strconv"
	"strings"
	"sync"
	"time"
	//"golang.org/x/text/unicode/norm"
)

var (
	TempDir = "./xl/worksheets/"
	rowPool = sync.Pool{
		New: func() interface{} {
			return &row{}
		},
	}
)

func CleanNonUtfAndControlChar(s string) string {
	s = strings.Map(func(r rune) rune {
		if r <= 31 { //if r is control character
			if r == 10 || r == 13 || r == 9 { //because newline
				return r
			}
			return -1
		}
		return r
	}, s)
	return s
}
func ExportWorksheet(filename string, rows RowFetcher, SharedStrWriter *bufio.Writer, cellsCount *int) error {
	destFile, err := os.Create(filename)
	if err != nil {
		return fmt.Errorf("Failed to create file: %s, error: %s", filename, err.Error())
	}
	defer destFile.Close()

	Writer := bufio.NewWriter(destFile)

	Writer.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">")
	Writer.WriteString("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"A1\" sqref=\"A1\"/></sheetView></sheetViews>")
	Writer.WriteString("<sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/>")
	Writer.WriteString("<sheetData>")

	rowCount := 1

	for {
		raw_row := rows.NextRow()
		if raw_row == nil {
			break
		}

		rr := rowPool.Get().(*row)
		rr.R = rowCount
		rr.C = []XlsxC{}

		for idx, val := range raw_row {
			colName := colCountToAlphaabet(idx)
			newCol := XlsxC{
				T: "s",
				R: fmt.Sprintf("%s%d", colName, rowCount),
				V: strconv.Itoa(*cellsCount),
			}
			*cellsCount++
			rr.C = append(rr.C, newCol)
			SharedStrWriter.WriteString(fmt.Sprintf("<si><t>%s</t></si>", html.EscapeString(CleanNonUtfAndControlChar(val))))
		}
		rr.Spans = "1:10"
		rr.Descent = "0.25"
		xmlBody, err := xml.Marshal(rr)
		rowPool.Put(rr)
		if err != nil {
			return err
		}
		_, err = Writer.Write(xmlBody)
		if err != nil {
			return err
		}

		if rowCount%1000 == 0 {
			SharedStrWriter.Flush()
			Writer.Flush()
		}
		rowCount++
	}
	Writer.WriteString("</sheetData>")
	Writer.WriteString("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>")
	Writer.WriteString("</worksheet>")
	return Writer.Flush()
}

func colCountToAlphaabet(idx int) string {
	var colName string
	if idx >= 26 {
		firstLetter := (idx / 26) - 1
		secondLetter := (idx % 26)
		colName = string(65+firstLetter) + string(65+secondLetter)
	} else {
		colName = string(65 + idx)
	}
	return strings.ToUpper(colName)
}
func Export(filename string, fetcher RowFetcher) error {
	now := time.Now()
	sheetName := now.Format("20060102150405") //filename should be (pseudo)random
	shaStr, err := os.Create(sheetName + ".ss")
	if err != nil {
		return err
	}
	defer func() {
		shaStr.Close()
		os.Remove("./" + sheetName + ".ss")
	}()

	SharedStrWriter := bufio.NewWriter(shaStr)
	SharedStrWriter.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
	SharedStrWriter.WriteString("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\">")
	cellCount := 0
	err = ExportWorksheet(sheetName, fetcher, SharedStrWriter, &cellCount)
	if err != nil {
		return err
	}

	SharedStrWriter.WriteString("</sst>")
	SharedStrWriter.Flush()
	outputFile := filename
	file := make(map[string]io.Reader)
	file["_rels/.rels"] = DummyRelsDotRels()
	file["docProps/app.xml"] = DummyAppXml()
	file["docProps/core.xml"] = DummyCoreXml()
	file["xl/_rels/workbook.xml.rels"] = DummyWorkbookRels()
	file["xl/theme/theme1.xml"] = DummyThemeXml()
	file["xl/worksheets/sheet1.xml"], _ = os.Open(sheetName)
	file["xl/styles.xml"] = DummyStyleXml()
	file["xl/workbook.xml"] = DummyWorkbookXml()
	file["xl/sharedStrings.xml"], _ = os.Open(sheetName + ".ss")
	file["[Content_Types].xml"] = DummyContentTypes()
	of, err := os.Create(outputFile)
	if err != nil {
		return fmt.Errorf("Failed to create file: %s, error: %s", outputFile, err.Error())
	}
	defer of.Close()
	zipWriter := zip.NewWriter(of)
	for k, v := range file {
		fWriter, err := zipWriter.Create(k)
		if err != nil {
			return err
		}
		if _, err = io.Copy(fWriter, v); err != nil {
			return err
		}
	}
	defer zipWriter.Close()
	(file["xl/sharedStrings.xml"].(*os.File)).Close()
	(file["xl/worksheets/sheet1.xml"].(*os.File)).Close()
	err = os.Remove("./" + sheetName)
	if err != nil {
		err = fmt.Errorf("Failed to remove file: %s, error: %s", sheetName, err.Error())
	}
	return err
}
