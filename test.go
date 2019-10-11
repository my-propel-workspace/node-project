package main

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// import "github.com/360EntSecGroup-Skylar/excelize"
func main() {
	f, err := excelize.OpenFile("./scm-alerts-report.xlsx")

	var libLocColName string
	var matchTypeColName string
	var sourceFileLocColName string

	// err = f.InsertCol("Alerts", "M")
	// f.Save()
	if err != nil {
		fmt.Println("error occured when opening file ", err.Error())
		return
	}
	alertsRows, err := f.GetRows("Alerts")

	if err != nil {
		fmt.Println("error occured when reading rows ", err)
		return
	}
	for alertsRowIndex, alertsRow := range alertsRows {
		// fmt.Println("rowsssssssssssssss length", len(rows), row)
		if alertsRowIndex == 0 {

			libLocColName, _ = excelize.ColumnNumberToName(len(alertsRow) + 1)
			fmt.Println("palese print this", libLocColName)

			matchTypeColName, _ = excelize.ColumnNumberToName(len(alertsRow) + 2)
			fmt.Println("palese print this", matchTypeColName)

			sourceFileLocColName, _ = excelize.ColumnNumberToName(len(alertsRow) + 3)
			fmt.Println("palese print this", sourceFileLocColName)

			categories := make(map[string]string)
			values := make(map[string]string)

			rowNumber := strconv.Itoa(alertsRowIndex + 1)

			categories[libLocColName+rowNumber] = "Library Location"
			categories[matchTypeColName+rowNumber] = "Match Type"
			categories[sourceFileLocColName+rowNumber] = "Source Files"

			for k, v := range categories {
				f.SetCellValue("Alerts", k, v)
			}

			// err := f.InsertCol("Alerts", "C")

			fmt.Println("cate and values are ", categories, values)
		}

		if alertsRowIndex != 0 {

			for n, colCell := range alertsRow {

				//compare alerts with library location
				libLocationFile, err := excelize.OpenFile("./scm-library-location.xlsx")

				if err != nil {
					fmt.Println("error occured when opening library location file ", err.Error(), n, colCell)
					return
				}
				libLocaRows, err := libLocationFile.GetRows("Library Location")

				if err != nil {
					fmt.Println("error occured when reading library location rows ", err)
					return
				}

				for libLocaRowIndex, libLocaRow := range libLocaRows {

					for libLocaColumnIndex, libLocaColumn := range libLocaRow {
						// fmt.Println("00000000000000000000000000000000000000", libLocaColumnIndex, libLocaRowIndex, libLocaColumn)

						if alertsRow[3] == libLocaRow[0] {

							libLocationVlues := make(map[string]string)

							rowNumber := strconv.Itoa(alertsRowIndex + 1)

							libLocationVlues[libLocColName+rowNumber] = libLocaRow[4]

							for k, v := range libLocationVlues {
								f.SetCellValue("Alerts", k, v)
							}

							fmt.Println("00000000000000000000000000000000000000", libLocaColumnIndex, libLocaRowIndex, libLocaColumn)
							break
						}

					}
				}

				//compare alerts with inventory report to set match type
				inventoryReportFile, err := excelize.OpenFile("./scm-inventory-report.xlsx")

				if err != nil {
					fmt.Println("error occured when opening inventory report file ", err.Error())
					return
				}
				inventoryReportRows, err := inventoryReportFile.GetRows("Inventory")

				if err != nil {
					fmt.Println("error occured when reading Inventory rows ", err)
					return
				}

				for inventoryReportRowIndex, inventoryReportRow := range inventoryReportRows {

					for libLocaColumnIndex, libLocaColumn := range inventoryReportRow {
						// fmt.Println("00000000000000000000000000000000000000", libLocaColumnIndex, inventoryReportRowIndex, libLocaColumn)

						if alertsRow[3] == inventoryReportRow[0] {

							inventoryReportValues := make(map[string]string)

							rowNumber := strconv.Itoa(alertsRowIndex + 1)

							inventoryReportValues[matchTypeColName+rowNumber] = inventoryReportRow[8]

							for k, v := range inventoryReportValues {
								f.SetCellValue("Alerts", k, v)
							}

							fmt.Println("00000000000000000000000000000000000000", libLocaColumnIndex, inventoryReportRowIndex, libLocaColumn)
							break
						}

					}
				}

				//compare alerts with source file location to set file locations
				sourceFileLocFile, err := excelize.OpenFile("./scm-source-file-inventory-report.xlsx")

				if err != nil {
					fmt.Println("error occured when opening source file location file ", err.Error())
					return
				}
				sourceFilesLocReportRows, err := sourceFileLocFile.GetRows("Source File Inventory")

				if err != nil {
					fmt.Println("error occured when reading Inventory rows ", err)
					return
				}
				var pathString = ""

				for sourceFilesLocReportRowIndex, sourceFilesLocReportRow := range sourceFilesLocReportRows {

					for sourceFilesLocReportColumnIndex, sourceFilesLocReportColumn := range sourceFilesLocReportRow {
						// fmt.Println("00000000000000000000000000000000000000", sourceFilesLocReportColumnIndex, sourceFilesLocReportRowIndex, sourceFilesLocReportColumn)
						if alertsRow[3] == sourceFilesLocReportRow[2] {
							pathString = pathString + sourceFilesLocReportRow[7] + ","
							fmt.Println("00000000000000000000000000000000000000", sourceFilesLocReportColumnIndex, sourceFilesLocReportRowIndex, sourceFilesLocReportColumn)
						}

						if sourceFilesLocReportRowIndex == len(sourceFilesLocReportRows)-1 {
							sourceFileLocReportValues := make(map[string]string)

							rowNumber := strconv.Itoa(alertsRowIndex + 1)

							sourceFileLocReportValues[sourceFileLocColName+rowNumber] = pathString

							for k, v := range sourceFileLocReportValues {
								f.SetCellValue("Alerts", k, v)
							}
						}

					}

				}

			}
		}
		// fmt.Println("+++++++++++++++++++++++++", m)
		// fmt.Println()
	}
	f.Save()
	fmt.Println("hello world")
}
