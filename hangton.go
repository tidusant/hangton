package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io/ioutil"
	"strings"
	"unicode"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/tidusant/c3m-common/log"

	//"io"

	"net/http"
	//	"os"
	"strconv"

	"github.com/gin-gonic/gin"
)

var mytoken string
var SheetName = "Sheet1"
var hangton []HangTonData
var errmsg = ""

func init() {
	hangton = []HangTonData{}
	getExcelData()
}

func main() {
	var port int
	var debug bool

	//fmt.Println(mycrypto.Encode("abc,efc", 5))
	flag.IntVar(&port, "port", 5084, "help message for flagname")
	flag.BoolVar(&debug, "debug", false, "Indicates if debug messages should be printed in log files")
	flag.StringVar(&mytoken, "token", ".fa1Xldsbe@", "Indicates if debug messages should be printed in log files")
	flag.Parse()

	//logLevel := log.DebugLevel
	if !debug {
		//logLevel = log.InfoLevel
		gin.SetMode(gin.ReleaseMode)
	}

	// log.SetOutputFile(fmt.Sprintf("portal-"+strconv.Itoa(port)), logLevel)
	// defer log.CloseOutputFile()
	// log.RedirectStdOut()

	log.Infof("running with port:" + strconv.Itoa(port))

	//init config

	router := gin.Default()

	router.GET("/hang/:search", func(c *gin.Context) {
		search := c.Param("search")
		search = strings.Trim(search, " ")
		strrt := ""

		if search != "" {
			search = strings.ToLower(search)
			strrt = searchhangton(search)

		} else {
			log.Debugf("check request error")
		}
		c.String(http.StatusOK, strrt)

	})

	router.POST("/hang/:search", func(c *gin.Context) {
		search := c.Param("search")
		search = strings.Trim(search, " ")
		strrt := ""

		if search != "" {
			search = strings.ToLower(search)
			strrt = searchhangton(search)

		} else {
			log.Debugf("check request error")
		}
		c.String(http.StatusOK, strrt)

	})

	router.GET("/file", func(c *gin.Context) {
		router.LoadHTMLGlob("html/*")

		c.HTML(http.StatusOK,
			// Use the index.html template
			"file.html",
			// Pass the data that the page uses (in this case, 'title')
			gin.H{
				"title": "Upload file",
			})

	})

	router.POST("/file", func(c *gin.Context) {
		router.LoadHTMLGlob("html/*")

		form, _ := c.MultipartForm()
		files := form.File["file"]
		for i, file := range files {
			log.Debugf("filename %d:%s - %s, %v", i, file.Filename, file.Header)

			filetmp, _ := file.Open()

			//file name

			filename := "tonkho.xlsx"
			data, err := ioutil.ReadAll(filetmp)
			if err != nil {
				errmsg = fmt.Sprintf("%v", err)
				c.String(http.StatusOK, errmsg)
			}
			err = ioutil.WriteFile("./data/"+filename, data, 0666)
			if err != nil {
				errmsg = fmt.Sprintf("%v", err)
				c.String(http.StatusOK, errmsg)
			}
		}

		c.HTML(http.StatusOK,
			// Use the index.html template
			"file.html",
			// Pass the data that the page uses (in this case, 'title')
			gin.H{
				"title": "Upload file",
			})

	})

	router.Run(":" + strconv.Itoa(port))

}

func searchhangton(search string) string {
	log.Debugf("searchhangton error message:%s", errmsg)
	if errmsg != "" {
		return errmsg
	}

	datareturns := []HangTonDataReturn{}
	for _, dat := range hangton {
		if strings.Index(dat.TenHang, search) >= 0 || strings.ToLower(dat.MaHang) == search {
			var d HangTonDataReturn
			d.MaHang = dat.MaHang
			d.TenHang = dat.TenHang
			d.Tong2Kho = dat.Tong2Kho
			d.UocLuongBan4Thang = dat.UocLuongBan4Thang
			d.TL1 = dat.TL1
			d.TL2 = dat.TL2
			d.TL3 = dat.TL3
			d.TL4 = dat.TL4
			d.TL5 = dat.TL5
			d.TL6 = dat.TL6
			datareturns = append(datareturns, d)
		}
	}
	b, _ := json.Marshal(datareturns)
	return string(b)
}

func getExcelData() {
	defer func() { //catch or finally
		if err := recover(); err != nil { //catch
			errmsg = fmt.Sprintf("Exception: %v", err)
		}
	}()

	xlsx, err := excelize.OpenFile("./data/tonkho.xlsx")
	if err != nil {
		errmsg = err.Error()

	}

	// Get all the rows in the Sheet1.
	rows := xlsx.GetRows(SheetName)

	sheetdata := xlsx.Sheet["xl/worksheets/sheet1.xml"]
	mergecells := sheetdata.MergeCells.Cells
	for irow, row := range rows {
		if irow < 5 {
			continue
		}
		var d HangTonData

		var rowdata []string
		for icol, colCell := range row {

			celldata := colCell
			if colCell == "" {
				for _, mergecell := range mergecells {
					ref := strings.Split(mergecell.Ref, ":")

					cellname := excelize.ToAlphaString(icol) + strconv.Itoa(irow+1)
					if checkCellInArea(cellname, mergecell.Ref) {
						celldata = xlsx.GetCellValue(SheetName, ref[0])
						//log.Debugf("getCellColRow %s: %s ", cellname, celldata)
						break
					}

					//log.Debugf("getCellColRow %s: %s %s")
					//fmt.Print(mergecell.Ref, "\t")
				}
			}
			rowdata = append(rowdata, celldata)
		}
		d.MaNhomHang1 = rowdata[0]
		d.MaNhomHang2 = rowdata[1]
		d.MaNhomHang3 = rowdata[2]
		d.MaNhomHang4 = rowdata[3]
		d.Kho = rowdata[4]
		d.MaHang = rowdata[5]
		d.TenHang = rowdata[6]
		d.Dvt = rowdata[7]
		d.TonCuoiSL, _ = strconv.Atoi(rowdata[24])
		d.Tong2Kho, _ = strconv.Atoi(rowdata[25])
		d.GiaHoreca, _ = strconv.Atoi(rowdata[26])
		d.BanTBThang, _ = strconv.Atoi(rowdata[27])
		d.UocLuongBan4Thang, _ = strconv.Atoi(rowdata[28])
		d.SLCanDauKy, _ = strconv.Atoi(rowdata[29])
		d.SLCanHienTai, _ = strconv.Atoi(rowdata[30])
		d.TL1, _ = strconv.Atoi(rowdata[31])
		d.TL2, _ = strconv.Atoi(rowdata[32])
		d.TL3, _ = strconv.Atoi(rowdata[33])
		d.TL4, _ = strconv.Atoi(rowdata[34])
		d.TL5, _ = strconv.Atoi(rowdata[35])
		d.TL6, _ = strconv.Atoi(rowdata[36])

		hangton = append(hangton, d)

	}
}

// checkCellInArea provides function to determine if a given coordinate is
// within an area.
func checkCellInArea(cell, area string) bool {
	cell = strings.ToUpper(cell)
	area = strings.ToUpper(area)

	ref := strings.Split(area, ":")
	if len(ref) < 2 {
		return false
	}

	from := ref[0]
	to := ref[1]

	col, row := getCellColRow(cell)
	fromCol, fromRow := getCellColRow(from)
	toCol, toRow := getCellColRow(to)

	return axisLowerOrEqualThan(fromCol, col) && axisLowerOrEqualThan(col, toCol) && axisLowerOrEqualThan(fromRow, row) && axisLowerOrEqualThan(row, toRow)
}

// axisLowerOrEqualThan returns true if axis1 <= axis2
// axis1/axis2 can be either a column or a row axis, e.g. "A", "AAE", "42", "1", etc.
//
// For instance, the following comparisons are all true:
//
// "A" <= "B"
// "A" <= "AA"
// "B" <= "AA"
// "BC" <= "ABCD" (in a XLSX sheet, the BC col comes before the ABCD col)
// "1" <= "2"
// "2" <= "11" (in a XLSX sheet, the row 2 comes before the row 11)
// and so on
func axisLowerOrEqualThan(axis1, axis2 string) bool {
	if len(axis1) < len(axis2) {
		return true
	} else if len(axis1) > len(axis2) {
		return false
	} else {
		return axis1 <= axis2
	}
}

// getCellColRow returns the two parts of a cell identifier (its col and row) as strings
//
// For instance:
//
// "C220" => "C", "220"
// "aaef42" => "aaef", "42"
// "" => "", ""
func getCellColRow(cell string) (col, row string) {
	for index, rune := range cell {
		if unicode.IsDigit(rune) {
			return cell[:index], cell[index:]
		}

	}

	return cell, ""
}
