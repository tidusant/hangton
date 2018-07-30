package main

import (
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
	"time"

	"github.com/gin-gonic/gin"
)

var mytoken string
var pagesize = 10
var SheetName = "Sheet1"
var hangton []HangTonData
var errmsg = ""
var updatetime = time.Now()
var uploadFilename = "data"

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
	flag.StringVar(&mytoken, "token", "485@XVNT", "Indicates if debug messages should be printed in log files")
	flag.IntVar(&pagesize, "pagesize", 10, "Indicates if debug messages should be printed in log files")
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
			strrt = searchhangton(search)

		} else {
			log.Debugf("check request error")
		}

		// c.Header("Content-Type", "application/json; charset=utf-8")
		// c.Next()
		// c.JSON(http.StatusOK, strrt)

		c.Data(200, "application/json; charset=utf-8", []byte(strrt))

		// log.Debugf("%s", strrt)
		// c.Bind(&hangton)
		// c.JSON(http.StatusOK, hangton)

		//c.String(http.StatusOK, strrt)

	})

	router.POST("/hang", func(c *gin.Context) {
		//search := c.Param("search")
		search := c.PostForm("text")
		//log.Debugf("search text %s %v", search, c.Params)

		strrt := ""

		if search != "" {
			strrt = searchhangton(search)

		} else {
			log.Debugf("check request error")
		}
		c.Header("Response-Type", "ephemeral")
		c.Header("Content-Type", "application/json")
		//https://api.slack.com/docs/message-attachments
		// c.String(http.StatusOK, strrt)

		c.Data(200, "application/json; charset=utf-8", []byte(strrt))

	})

	router.POST("/tonkho", func(c *gin.Context) {
		//search := c.Param("search")
		search := c.PostForm("text")
		//log.Debugf("search text %s %v", search, c.Params)

		strrt := ""

		if search == mytoken {
			strrt = searchhangton(search)

		} else {
			log.Debugf("check request error")
		}
		c.Header("Response-Type", "ephemeral")
		c.Header("Content-Type", "application/json")
		//https://api.slack.com/docs/message-attachments
		// c.String(http.StatusOK, strrt)

		c.Data(200, "application/json; charset=utf-8", []byte(strrt))

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
			uploadFilename = file.Filename
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
		getExcelData()
		message := "Done, " + strconv.Itoa(len(hangton)) + " rows were updated"
		if errmsg != "" {
			message = errmsg
		}
		c.HTML(http.StatusOK,
			// Use the index.html template
			"file.html",
			// Pass the data that the page uses (in this case, 'title')
			gin.H{
				"title":   "Upload file",
				"message": message,
			})

	})

	router.Run(":" + strconv.Itoa(port))

}

func searchhangton(search string) string {
	//log.Debugf("searchhangton error message:%s", errmsg)
	if errmsg != "" {
		return errmsg
	}

	searches := strings.Split(search, " ")
	isAuth := false
	isTonKho := false
	page := 1
	//check page
	if len(searches) > 1 && searches[len(searches)-1][:4] == "page" {
		page, _ = strconv.Atoi(searches[len(searches)-1][4:])
		searches = searches[:len(searches)-1]
		search = strings.Join(searches, " ")
	}
	if search == mytoken {
		isTonKho = true
		isAuth = true
		search = ""
	} else if searches[len(searches)-1] == mytoken {
		isAuth = true
		searches = searches[:len(searches)-1]
		search = strings.Join(searches, " ")
	}
	search = strings.Trim(strings.ToLower(search), " ")

	data := ``
	count := 0
	text := ``
	datahang := make(map[string]HangTonData)
	outcount := 0
	matchcount := 0
	for _, dat := range hangton {

		if (isTonKho && dat.SLCanHienTai < 0) || strings.Index(strings.ToLower(dat.TenHang), search) >= 0 || strings.ToLower(dat.MaHang) == search {
			//check paging
			matchcount++
			if matchcount-1 < (page-1)*pagesize {
				continue
			}
			if outcount > pagesize-1 {
				break
			}
			//check exist
			if _, ok := datahang[dat.MaHang]; ok {
				dattemp := datahang[dat.MaHang]
				dattemp.UocLuongBan4Thang = dat.UocLuongBan4Thang
				dattemp.GiaHoreca += " " + dat.GiaHoreca
				for key, sl := range dat.TL {
					dattemp.TL[key] += " " + sl

				}
				datahang[dat.MaHang] = dattemp
			} else {
				datahang[dat.MaHang] = dat
				outcount++
			}

		}

	}
	for _, dat := range datahang {
		color := "#7CD197"
		if count%2 == 0 {
			color = "#F35A00"
		}
		//arrival
		arrival := ""
		for name, sl := range dat.TL {
			if len(strings.Trim(sl, " ")) > 0 {
				arrival += `,{"title": "` + name + `",
				"value": "` + sl + `",
				"short": false}`
			}
		}

		data += `{`
		data += `"title": "` + dat.TenHang + `",
			"title_link": "https://phuem.com/",
			"color": "` + color + `",
			"fields": [
                {
                    "title": "Mã Hàng",
                    "value": "` + dat.MaHang + `",
                    "short": true
                },
                {
                    "title": "Tổng 2 Kho",
                    "value": "` + strconv.Itoa(dat.Tong2Kho) + `",
                    "short": true
				}`
		data += `,
                {
                    "title": "Ước Lượng Bán 4 tháng",
                    "value": "` + strconv.Itoa(dat.UocLuongBan4Thang) + `",
                    "short": false
				}`
		data += `,
                {
                    "title": "Giá Horeca",
                    "value": "` + dat.GiaHoreca + `",
                    "short": false
				}`
		if isAuth {

			data += `,
                {
                    "title": "SL cần hiện tại",
                    "value": "` + strconv.Itoa(dat.SLCanHienTai) + `",
                    "short": false
				}`
		}
		if isAuth {
			data += `,
                {
                    "title": "SL cần đầu kỳ",
                    "value": "` + strconv.Itoa(dat.SLCanDauKy) + `",
                    "short": false
				}`
		}

		data += arrival + `]},`
		count++
	}
	//show paging
	if matchcount > page*pagesize {
		strnextpage := "/hang"
		if isTonKho {
			strnextpage = "/tonkho"
		}
		strnextpage += " " + search
		if isAuth {
			strnextpage += " " + mytoken
		}
		strnextpage += " page" + strconv.Itoa(page+1)
		data += `{"title":"Next page: ` + strnextpage + `"}`
	} else {
		data = data[:len(data)-1]
	}

	text += `{"text":"`
	attachments := ""
	if count > 0 {
		text += strconv.Itoa(count) + ` founds\n`

		attachments = `,"attachments": [` + data + `]`
	} else {
		text += ` not founds\n`
	}

	text += `*` + uploadFilename + `* updated at: ` + updatetime.Format("15:04 02-01-2006") + `" ` + attachments + `}`

	return text
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
	hdata := []HangTonData{}
	for irow, row := range rows {
		if irow < 5 {
			continue
		}
		var d HangTonData
		d.TL = make(map[string]string)
		var rowdata []string
		for icol, colCell := range row {

			celldata := colCell
			//check mergcell
			if colCell == "" {
				for _, mergecell := range mergecells {
					ref := strings.Split(mergecell.Ref, ":")

					cellname := excelize.ToAlphaString(icol) + strconv.Itoa(irow+1)
					if checkCellInArea(cellname, mergecell.Ref) {
						celldata = xlsx.GetCellValue(SheetName, ref[0])
						break
					}
				}
			}
			//check name column
			colname := xlsx.GetCellValue(SheetName, excelize.ToAlphaString(icol)+"4")
			colname = strings.Replace(strings.Replace(strings.Replace(colname, "\r\n", " ", -1), "\n", " ", -1), "\"", "“", -1)
			colnametrim := strings.Trim(strings.ToLower(colname), " ")
			celldata = strings.Replace(strings.Replace(strings.Replace(celldata, "\r\n", " ", -1), "\n", " ", -1), "\"", "“", -1)
			if colnametrim == "mã nhóm hàng 1" {
				d.MaNhomHang1 = celldata
			} else if colnametrim == "mã nhóm hàng 2" {
				d.MaNhomHang2 = celldata
			} else if colnametrim == "mã nhóm hàng 3" {
				d.MaNhomHang3 = celldata
			} else if colnametrim == "mã nhóm hàng 4" {
				d.MaNhomHang4 = celldata
			} else if colnametrim == "kho" {
				d.Kho = celldata
			} else if colnametrim == "mã hàng" {
				d.MaHang = celldata
			} else if colnametrim == "tên hàng" {
				d.TenHang = celldata
			} else if colnametrim == "đtv" {
				d.Dvt = celldata
			} else if colnametrim == "tồn cuối sl" {
				d.TonCuoiSL, _ = strconv.Atoi(celldata)
			} else if colnametrim == "tổng 2 kho" {
				d.Tong2Kho, _ = strconv.Atoi(celldata)
			} else if colnametrim == "giá horeca" {
				d.GiaHoreca = celldata
			} else if colnametrim == "bán trung bình tháng" {
				d.BanTBThang, _ = strconv.Atoi(celldata)
			} else if colnametrim == "ước lượng bán 4 tháng" {
				d.UocLuongBan4Thang, _ = strconv.Atoi(celldata)
			} else if colnametrim == "số lượng cần đầu kỳ" {
				d.SLCanDauKy, _ = strconv.Atoi(celldata)
			} else if colnametrim == "số lượng cần hiện tại" {
				d.SLCanHienTai, _ = strconv.Atoi(celldata)
			} else if len(colname) > 8 && strings.ToLower(colname[:9]) == "arriving-" {
				d.TL[colname] = celldata
			}
			rowdata = append(rowdata, celldata)
		}

		hdata = append(hdata, d)

	}
	hangton = hdata
	errmsg = ""
	updatetime = time.Now()
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
