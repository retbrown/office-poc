package main

import (
	"fmt"
	"log"

	"github.com/unidoc/unioffice/chart"
	"github.com/unidoc/unioffice/color"
	"github.com/unidoc/unioffice/common/license"
	"github.com/unidoc/unioffice/document"
	"github.com/unidoc/unioffice/measurement"
	"github.com/unidoc/unioffice/schema/soo/wml"
	"github.com/unidoc/unioffice/spreadsheet"
)

const licenseKey = `
-----BEGIN UNIDOC LICENSE KEY-----
eyJsaWNlbnNlX2lkIjoiMjQ0OWYxOGItNGQ3Mi00ODFhLTQyMmYtZjM0MGMyOGIzMDE0IiwiY3VzdG9tZXJfaWQiOiI3OGVmY2IxZi0xYTRlLTQ0NmUtNjNlNi0yOTgwNDMzZTQ2NTAiLCJjdXN0b21lcl9uYW1lIjoiUkVhbHlzZSIsImN1c3RvbWVyX2VtYWlsIjoicm9iZXJ0QHJlYWx5c2UuY29tIiwidGllciI6ImJ1c2luZXNzIiwiY3JlYXRlZF9hdCI6MTYwNzM1ODkyOCwiZXhwaXJlc19hdCI6MTYwODU5NTE5OSwiY3JlYXRvcl9uYW1lIjoiVW5pRG9jIFN1cHBvcnQiLCJjcmVhdG9yX2VtYWlsIjoic3VwcG9ydEB1bmlkb2MuaW8iLCJ1bmlwZGYiOnRydWUsInVuaW9mZmljZSI6dHJ1ZSwidHJpYWwiOnRydWV9
+
cMEDZnG6SpTU8zRlg2vsKrCZYYd8GWazxGxZU/W4lahOamFxUBU6Qg/zVDCENGFUKMmNr05uKDlV8Sun+PGM/bAq22EN9RU1WpHZ4ruddEHP/9oSumF+GdfuEkhVhhNsmjPagPTDi1Fa9bEAVLkSIyCC1xy2Z5W085kFuV40VWP3dpFj0lG8CyUcQg2KRPGGljbcyWi1c5YMZOAriIg+VO7BoEvK5AJTeA8BTawcSof2O/sHTm2YkE1mdp+WyPBnfQONTn1xD3K52ncEEXOTp3Orm01lICxQhrdYTcnY2wF+X4SWYhczPtQ0JqzL4o0FiEpDXJHdhSbdbmEDXvs9Tg==
-----END UNIDOC LICENSE KEY-----
`

func init() {
	err := license.SetLicenseKey(licenseKey, `REalyse`)
	if err != nil {
		panic(err)
	}
}

func main() {
	doc := document.New()
	defer doc.Close()

	hdr := doc.AddHeader()

	para := hdr.AddParagraph()
	para.Properties().AddTabStop(2.5*measurement.Inch, wml.ST_TabJcCenter, wml.ST_TabTlcNone)
	run := para.AddRun()
	run.AddTab()
	run.AddText("This is a header")

	// Headers and footers are not immediately associated with a document as a
	// document can have multiple headers and footers for different sections.
	doc.BodySection().SetHeader(hdr, wml.ST_HdrFtrDefault)

	ftr := doc.AddFooter()
	para = ftr.AddParagraph()
	para.Properties().AddTabStop(6*measurement.Inch, wml.ST_TabJcRight, wml.ST_TabTlcNone)
	run = para.AddRun()
	run.AddText("This is my footer")
	run.AddTab()
	run.AddText("Pg ")
	run.AddField(document.FieldCurrentPage)
	run.AddText(" of ")
	run.AddField(document.FieldNumberOfPages)
	doc.BodySection().SetFooter(ftr, wml.ST_HdrFtrDefault)

	// First Table
	{
		table := doc.AddTable()
		// width of the page
		table.Properties().SetWidthPercent(100)
		// with thick borers
		borders := table.Properties().Borders()
		borders.SetAll(wml.ST_BorderSingle, color.Auto, 2*measurement.Point)

		row := table.AddRow()
		run := row.AddCell().AddParagraph().AddRun()
		run.AddText("Name")
		run.Properties().SetHighlight(wml.ST_HighlightColorYellow)
		row.AddCell().AddParagraph().AddRun().AddText("John Smith")
		row = table.AddRow()
		row.AddCell().AddParagraph().AddRun().AddText("Street Address")
		row.AddCell().AddParagraph().AddRun().AddText("111 Country Road")
	}

	doc.AddParagraph() // break up the consecutive tables

	// Second Table
	{
		table := doc.AddTable()
		// 4 inches wide
		table.Properties().SetWidth(4 * measurement.Inch)
		borders := table.Properties().Borders()
		// thin borders
		borders.SetAll(wml.ST_BorderSingle, color.Auto, measurement.Zero)

		row := table.AddRow()
		cell := row.AddCell()
		// column span / merged cells
		cell.Properties().SetColumnSpan(2)

		run := cell.AddParagraph().AddRun()
		run.AddText("Cells can span multiple columns")

		row = table.AddRow()
		cell = row.AddCell()
		cell.Properties().SetVerticalMerge(wml.ST_MergeRestart)
		cell.AddParagraph().AddRun().AddText("Vertical Merge")
		row.AddCell().AddParagraph().AddRun().AddText("")

		row = table.AddRow()
		cell = row.AddCell()
		cell.Properties().SetVerticalMerge(wml.ST_MergeContinue)
		cell.AddParagraph().AddRun().AddText("Vertical Merge 2")
		row.AddCell().AddParagraph().AddRun().AddText("")

		row = table.AddRow()
		row.AddCell().AddParagraph().AddRun().AddText("Street Address")
		row.AddCell().AddParagraph().AddRun().AddText("111 Country Road")
	}

	doc.AddParagraph()
	if err := doc.Validate(); err != nil {
		log.Fatalf("error during validation: %s", err)
	}

	doc.SaveToFile("header-footer.docx")

	ss := spreadsheet.New()
	defer ss.Close()
	sheet := ss.AddSheet()

	// Create all of our data
	row := sheet.AddRow()
	row.AddCell().SetString("Item")
	row.AddCell().SetString("Price")
	row.AddCell().SetString("# Sold")
	row.AddCell().SetString("Total")
	for r := 0; r < 5; r++ {
		row := sheet.AddRow()
		row.AddCell().SetString(fmt.Sprintf("Product %d", r+1))
		row.AddCell().SetNumber(1.23 * float64(r+1))
		row.AddCell().SetNumber(float64(r%3 + 1))
		row.AddCell().SetFormulaRaw(fmt.Sprintf("C%d*B%d", r+2, r+2))
	}

	// Charts need to reside in a drawing
	dwng := ss.AddDrawing()
	chrt1, anc1 := dwng.AddChart(spreadsheet.AnchorTypeTwoCell)
	chrt2, anc2 := dwng.AddChart(spreadsheet.AnchorTypeTwoCell)
	addBarChart(chrt1)
	addLineChart(chrt2)
	anc1.SetWidth(9)
	anc1.MoveTo(5, 1)
	anc2.MoveTo(1, 23)

	// and finally add the chart to the sheet
	sheet.SetDrawing(dwng)

	if err := ss.Validate(); err != nil {
		log.Fatalf("error validating sheet: %s", err)
	}
	ss.SaveToFile("multiple-chart.xlsx")
}

func addBarChart(chrt chart.Chart) {
	chrt.AddTitle().SetText("Bar Chart")
	lc := chrt.AddBarChart()
	priceSeries := lc.AddSeries()
	priceSeries.SetText("Price")
	// Set a category axis reference on the first series to pull the product names
	priceSeries.CategoryAxis().SetLabelReference(`'Sheet 1'!A2:A6`)
	priceSeries.Values().SetReference(`'Sheet 1'!B2:B6`)

	soldSeries := lc.AddSeries()
	soldSeries.SetText("Sold")
	soldSeries.Values().SetReference(`'Sheet 1'!C2:C6`)

	totalSeries := lc.AddSeries()
	totalSeries.SetText("Total")
	totalSeries.Values().SetReference(`'Sheet 1'!D2:D6`)

	// the line chart accepts up to two axes
	ca := chrt.AddCategoryAxis()
	va := chrt.AddValueAxis()
	lc.AddAxis(ca)
	lc.AddAxis(va)

	ca.SetCrosses(va)
	va.SetCrosses(ca)
}

func addLineChart(chrt chart.Chart) {
	chrt.AddTitle().SetText("Line Chart")
	lc := chrt.AddLineChart()
	priceSeries := lc.AddSeries()
	priceSeries.SetText("Price")
	// Set a category axis reference on the first series to pull the product names
	priceSeries.CategoryAxis().SetLabelReference(`'Sheet 1'!A2:A6`)
	priceSeries.Values().SetReference(`'Sheet 1'!B2:B6`)

	soldSeries := lc.AddSeries()
	soldSeries.SetText("Sold")
	soldSeries.Values().SetReference(`'Sheet 1'!C2:C6`)

	totalSeries := lc.AddSeries()
	totalSeries.SetText("Total")
	totalSeries.Values().SetReference(`'Sheet 1'!D2:D6`)

	// the line chart accepts up to two axes
	ca := chrt.AddCategoryAxis()
	va := chrt.AddValueAxis()
	lc.AddAxis(ca)
	lc.AddAxis(va)

	ca.SetCrosses(va)
	va.SetCrosses(ca)
}
