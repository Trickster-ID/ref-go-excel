package main

import (
	"fmt"
	"mime/multipart"
	"net/http"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/gin-gonic/gin"
	"github.com/tealeg/xlsx"
)

type TFData struct {
	OriginFieldValue      string `json:"origin_field_value"`
	DestinationFieldValue string `json:"destination_field_value"`
}

type DataDTO struct {
	ClientID  string                `json:"client_id" form:"client_id"`
	ExcelFile *multipart.FileHeader `json:"excel_file" form:"excel_file"`
}

func main() {

	r := gin.Default()
	r.POST("/import", readExcel)
	r.GET("/export", createExcel)
	r.Run("127.0.0.1:8888")
}

// convert multipart.fileheader xlsx file into struct
func readExcel(c *gin.Context) {
	var dto DataDTO
	err := c.ShouldBind(&dto)
	fmt.Println(dto)
	if err != nil {
		c.String(http.StatusBadRequest, fmt.Sprintf("error shouldbind : %s", err.Error()))
		return
	}

	openedFile, err := dto.ExcelFile.Open()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}
	defer openedFile.Close()

	xlFile, err := xlsx.OpenReaderAt(openedFile, dto.ExcelFile.Size)

	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}

	sheet := xlFile.Sheets[0]

	var tfDatas []TFData

	for i := 1; i < len(sheet.Rows); i++ {
		tfData := TFData{
			OriginFieldValue:      sheet.Rows[i].Cells[0].Value,
			DestinationFieldValue: sheet.Rows[i].Cells[1].Value,
		}
		tfDatas = append(tfDatas, tfData)
	}

	fmt.Println("berhasil")
	fmt.Println("data : ", tfDatas)

	c.JSON(http.StatusOK, gin.H{
		"client_id": dto.ClientID,
		"message":   tfDatas,
	})
}

func createExcel(c *gin.Context) {
	data := []TFData{
		{
			OriginFieldValue:      "row 1 column 1",
			DestinationFieldValue: "row 1 column 2",
		},
		{
			OriginFieldValue:      "row 2 column 1",
			DestinationFieldValue: "row 2 column 2",
		},
	}

	newFile := excelize.NewFile()

	index := newFile.NewSheet("Sheet1")

	newFile.SetCellValue("Sheet1", "A1", "Origin Field Value")
	newFile.SetCellValue("Sheet1", "B1", "Destination Field Value")

	for i, v := range data {
		newFile.SetCellValue("Sheet1", fmt.Sprintf("A%d", i+2), v.OriginFieldValue)
		newFile.SetCellValue("Sheet1", fmt.Sprintf("B%d", i+2), v.DestinationFieldValue)
	}

	newFile.SetActiveSheet(index)

	newFile.SaveAs("pertamacoba.xlsx")

	c.JSON(http.StatusOK, gin.H{
		"client_id": "test",
		"message":   "mantul",
	})
}
