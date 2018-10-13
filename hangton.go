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
var hangtonsg []HangTonData
var hangtondn []HangTonData
var hangtonhn []HangTonData
var errmsg = ""
var errmsgdn = ""
var errmsgsg = ""
var errmsghn = ""
var updatetime = time.Now()
var uploadFilename = "tonkho"
var updatetimedn = time.Now()
var uploadFilenamedn = "tonkhodn"
var updatetimehn = time.Now()
var uploadFilenamehn = "tonkhodn"

type xlsxMergeCell struct {
	Ref string `xml:"ref,attr,omitempty"`
}

func init() {
	hangtonsg = []HangTonData{}
	getExcelData("sg")
	errmsgsg = errmsg
	hangtondn = []HangTonData{}
	getExcelData("dn")
	errmsgdn = errmsg
	hangtonhn = []HangTonData{}
	getExcelData("hn")
	errmsghn = errmsg
}

func main() {
	var port int
	var debug bool

	//fmt.Println(mycrypto.Encode("abc,efc", 5))
	flag.IntVar(&port, "port", 5084, "help message for flagname")
	flag.BoolVar(&debug, "debug", false, "Indicates if debug messages should be printed in log files")
	flag.StringVar(&mytoken, "token", "489xvnt", "Indicates if debug messages should be printed in log files")
	flag.IntVar(&pagesize, "pagesize", 30, "Indicates if debug messages should be printed in log files")
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
			strrt = searchhangton(search, "sg")
		} else {
			log.Debugf("check request error")
		}
		c.Data(200, "application/json; charset=utf-8", []byte(strrt))
	})
	router.GET("/hangdn/:search", func(c *gin.Context) {
		search := c.Param("search")
		search = strings.Trim(search, " ")
		strrt := ""

		if search != "" {
			strrt = searchhangton(search, "dn")
		} else {
			log.Debugf("check request error")
		}
		c.Data(200, "application/json; charset=utf-8", []byte(strrt))
	})

	router.GET("/hanghn/:search", func(c *gin.Context) {
		search := c.Param("search")
		search = strings.Trim(search, " ")
		strrt := ""

		if search != "" {
			strrt = searchhangton(search, "hn")
		} else {
			log.Debugf("check request error")
		}
		c.Data(200, "application/json; charset=utf-8", []byte(strrt))
	})

	router.POST("/hang", func(c *gin.Context) {
		//search := c.Param("search")
		search := c.PostForm("text")
		//log.Debugf("search text %s %v", search, c.Params)

		strrt := ""

		if search != "" {
			strrt = searchhangton(search, "sg")

		} else {
			log.Debugf("check request error")
		}
		c.Header("Response-Type", "ephemeral")
		c.Header("Content-Type", "application/json")

		c.Data(200, "application/json; charset=utf-8", []byte(strrt))

	})

	router.POST("/hangdn", func(c *gin.Context) {
		//search := c.Param("search")
		search := c.PostForm("text")
		//log.Debugf("search text %s %v", search, c.Params)

		strrt := ""

		if search != "" {
			strrt = searchhangton(search, "dn")

		} else {
			log.Debugf("check request error")
		}
		c.Header("Response-Type", "ephemeral")
		c.Header("Content-Type", "application/json")

		c.Data(200, "application/json; charset=utf-8", []byte(strrt))

	})

	router.POST("/hanghn", func(c *gin.Context) {
		//search := c.Param("search")
		search := c.PostForm("text")
		//log.Debugf("search text %s %v", search, c.Params)

		strrt := ""

		if search != "" {
			strrt = searchhangton(search, "hn")

		} else {
			log.Debugf("check request error")
		}
		c.Header("Response-Type", "ephemeral")
		c.Header("Content-Type", "application/json")

		c.Data(200, "application/json; charset=utf-8", []byte(strrt))

	})

	router.POST("/tonkho", func(c *gin.Context) {
		//search := c.Param("search")
		search := c.PostForm("text")
		//log.Debugf("search text %s %v", search, c.Params)

		strrt := ""

		if search != "" {
			strrt = searchhangton(search, "sg")

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
		filetype := c.PostForm("filetype")
		errmsg = ""
		fileuploadname := ""
		for i, file := range files {
			log.Debugf("filename %d:%s - %s, %v", i, file.Filename, file.Header)

			filetmp, _ := file.Open()

			//file name
			fileuploadname = file.Filename
			log.Debugf("filename %d:%s - %s", i, fileuploadname[len(fileuploadname)-5:])
			if fileuploadname[len(fileuploadname)-5:] != ".xlsx" {
				errmsg = "ERROR: must upload .xlsx file extension!"
			} else {
				filename := "tonkho.xlsx"
				if filetype == "dn" {
					filename = "tonkhodn.xlsx"
				} else if filetype == "hn" {
					filename = "tonkhohn.xlsx"
				}
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

		}
		if errmsg == "" {
			getExcelData(filetype)
		}

		message := ""
		if errmsg == "" {
			if filetype == "dn" {
				errmsgdn = ""
				message = "Done, " + strconv.Itoa(len(hangtondn)) + " rows were updated"
				uploadFilenamedn = fileuploadname
			} else if filetype == "hn" {
				errmsgdn = ""
				message = "Done, " + strconv.Itoa(len(hangtonhn)) + " rows were updated"
				uploadFilenamehn = fileuploadname
			} else {
				errmsgsg = ""
				message = "Done, " + strconv.Itoa(len(hangtonsg)) + " rows were updated"
				uploadFilename = fileuploadname
			}

		} else {
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

func searchhangton(search, filetype string) string {
	//log.Debugf("searchhangton error message:%s", errmsg)
	hangtonsearch := []HangTonData{}
	if filetype == "sg" {
		if errmsgsg != "" {
			return errmsgsg
		}
		hangtonsearch = hangtonsg
	} else if filetype == "dn" {
		if errmsgdn != "" {
			return errmsgdn
		}
		hangtonsearch = hangtondn
	} else if filetype == "hn" {
		if errmsghn != "" {
			return errmsghn
		}
		hangtonsearch = hangtonhn
	}

	searches := strings.Split(search, " ")
	isAuth := false
	isTonKho := false
	page := 1
	//check page
	if len(searches) > 1 && searches[len(searches)-1][:1] == "p" {
		page, _ = strconv.Atoi(searches[len(searches)-1][1:])
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
	//search words
	wordsearchs := strings.Split(search, " ")

	data := ``
	count := 0
	text := ``
	datamatch := make(map[string]HangTonData)
	datamatch2 := make(map[string]HangTonData)
	dataref := []string{}
	dataref2 := []string{}
	//outcount := 0
	//matchcount := 0
	for _, dat := range hangtonsearch {
		isMatch := false
		if isTonKho && dat.SLCanHienTai < 0 {
			//log.Debugf("matched: %v", dat.SLCanHienTai)
			isMatch = true
		} else {

			//loop word search
			for _, word := range wordsearchs {
				if word != "" && (strings.Index(strings.ToLower(dat.TenHang), word) >= 0 || strings.ToLower(dat.MaHang) == word) {
					isMatch = true
					//log.Debugf("fail matched: %s - %s - %s", dat.TenHang, word, search)
					break
				}
			}
		}

		if isMatch {
			//log.Debugf("matched: %v", dat)
			// //check paging
			// matchcount++
			// if matchcount-1 < (page-1)*pagesize {
			// 	continue
			// }
			// if outcount > pagesize-1 {
			// 	break
			// }

			//exactly match
			if isTonKho || strings.Index(strings.ToLower(dat.TenHang), search) >= 0 || strings.ToLower(dat.MaHang) == search {
				//check exist
				log.Debugf("exactly matched: %v", dat.SLCanHienTai)
				if _, ok := datamatch[dat.MaHang]; ok {

					dattemp := datamatch[dat.MaHang]
					dattemp.UocLuongBan4Thang = dat.UocLuongBan4Thang

					dattemp.Tong2Kho += dat.TonCuoiSL
					dattemp.GiaHoreca += " " + dat.GiaHoreca
					for key, sl := range dat.TL {
						dattemp.TL[key] += " " + sl

					}
					datamatch[dat.MaHang] = dattemp
				} else {
					dat.Tong2Kho = dat.TonCuoiSL
					datm := dat
					datm.TL = make(map[string]string)
					datamatch[dat.MaHang] = datm
					//deep copy
					for key, val := range dat.TL {
						datamatch[dat.MaHang].TL[key] = val
					}

					//datamatch[dat.MaHang].TL = c3mcommon.CopyMap(dat.TL)
					dataref = append(dataref, dat.MaHang)
					//outcount++
				}
				//log.Debugf("exactly matched: %v", dat.MaHang)
			} else {
				//check exist
				//log.Debugf("check exist %v", datamatch2[dat.MaHang])
				if _, ok := datamatch2[dat.MaHang]; ok {
					//log.Debugf("exist: %v", dat.MaHang)
					dattemp := datamatch2[dat.MaHang]
					dattemp.UocLuongBan4Thang = dat.UocLuongBan4Thang
					dattemp.Tong2Kho += dattemp.TonCuoiSL
					dattemp.GiaHoreca += " " + dat.GiaHoreca
					for key, sl := range dat.TL {
						dattemp.TL[key] += " " + sl

					}
					datamatch2[dat.MaHang] = dattemp
				} else {
					dat.Tong2Kho = dat.TonCuoiSL
					datm := dat
					datm.TL = make(map[string]string)
					datamatch2[dat.MaHang] = datm
					//deep copy
					for key, val := range dat.TL {
						datamatch2[dat.MaHang].TL[key] = val
					}
					dataref2 = append(dataref2, dat.MaHang)
					//outcount++
				}
				//log.Debugf("partial matched: %v", dataref2)
			}

		}

	}

	//reorder to get match exactly first
	var datashow []HangTonData
	for _, mahang := range dataref {
		datashow = append(datashow, datamatch[mahang])

	}
	if len(datashow) == 0 {
		for _, mahang := range dataref2 {
			datashow = append(datashow, datamatch2[mahang])
		}
	}
	//log.Debugf("datamatch: %v", datamatch)
	for i, dat := range datashow {
		//check paging

		if i < (page-1)*pagesize {
			continue
		}
		if i >= page*pagesize {
			break
		}

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
                    "title": "Tổng Kho",
                    "value": "` + strconv.Itoa(dat.Tong2Kho) + `",
                    "short": true
				}`
		if filetype == "sg" {
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
		}
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
	if count+pagesize*(page-1) < len(datashow) {
		strnextpage := "/hang " + search
		if filetype == "dn" || filetype == "hn" {
			strnextpage = "/hang" + filetype + " " + search
		}
		if isTonKho {
			strnextpage = "/tonkho"
		}
		if isAuth {
			strnextpage += " " + mytoken
		}
		strnextpage += " p" + strconv.Itoa(page+1)
		data += `{"title":"Next page: ` + strnextpage + `"}`
	} else {
		if len(data) > 0 {
			data = data[:len(data)-1]
		}
	}

	text += `{"text":"`
	attachments := ""
	if count > 0 {
		text += strconv.Itoa(count) + ` founds\n`

		attachments = `,"attachments": [` + data + `]`
	} else {
		text += ` not founds\n`
	}
	filename := uploadFilename
	updatetimestr := updatetime.Format("15:04 02-01-2006")
	if filetype == "dn" {
		filename = uploadFilenamedn
		updatetimestr = updatetimedn.Format("15:04 02-01-2006")
	} else if filetype == "hn" {
		filename = uploadFilenamehn
		updatetimestr = updatetimehn.Format("15:04 02-01-2006")
	}
	text += `*` + filename + `* updated at: ` + updatetimestr + `" ` + attachments + `}`

	return text
}

func getExcelData(filetype string) {
	errmsg = ""
	defer func() { //catch or finally
		if err := recover(); err != nil { //catch
			errmsg = fmt.Sprintf("Exception: %v", err)
		}
	}()
	filename := "tonkho.xlsx"
	colRequire := make(map[string]bool)
	if filetype == "dn" {
		filename = "tonkhodn.xlsx"
		colRequire = map[string]bool{
			"mã hàng":     false,
			"tên hàng":    false,
			"tồn cuối sl": false,
			"giá horeca":  false,
		}
	} else if filetype == "hn" {
		filename = "tonkhohn.xlsx"
		colRequire = map[string]bool{
			"mã hàng":     false,
			"tên hàng":    false,
			"tồn cuối sl": false,
			"giá horeca":  false,
		}
	} else {
		colRequire = map[string]bool{
			"mã hàng":               false,
			"tên hàng":              false,
			"tồn cuối sl":           false,
			"giá horeca":            false,
			"ước lượng bán 4 tháng": false,
			"số lượng cần hiện tại": false,
			"số lượng cần đầu kỳ":   false,
		}
	}
	xlsx, err := excelize.OpenFile("./data/" + filename)
	if err != nil {
		errmsg = err.Error()

	}

	// Get all the rows in the Sheet1.
	rows := xlsx.GetRows(xlsx.GetSheetName(xlsx.GetActiveSheetIndex()))

	sheetdata := xlsx.Sheet["xl/worksheets/sheet1.xml"]
	var mergecellRef []string
	if sheetdata.MergeCells != nil {
		for _, cell := range sheetdata.MergeCells.Cells {
			mergecellRef = append(mergecellRef, cell.Ref)
		}
	}

	hdata := []HangTonData{}
	//check column require
	headerrow := 1
	if filetype == "sg" {
		headerrow = 4
	}
	for icol, _ := range rows[0] {
		colname := xlsx.GetCellValue(SheetName, excelize.ToAlphaString(icol)+strconv.Itoa(headerrow))
		colname = strings.Replace(strings.Replace(strings.Replace(colname, "\r\n", " ", -1), "\n", " ", -1), "\"", "“", -1)
		colnametrim := strings.Trim(strings.ToLower(colname), " ")
		if _, ok := colRequire[colnametrim]; ok {
			colRequire[colnametrim] = true
		}
	}
	for colname, colreq := range colRequire {
		if !colreq {
			errmsg = `ERROR! Column "` + colname + `" is required!`
			return
		}
	}

	for irow, row := range rows {
		if irow <= headerrow-1 {
			continue
		}
		var d HangTonData
		d.TL = make(map[string]string)
		var rowdata []string

		for icol, colCell := range row {

			celldata := colCell
			//check mergcell
			if colCell == "" && len(mergecellRef) > 0 {

				for _, cellRef := range mergecellRef {
					ref := strings.Split(cellRef, ":")

					cellname := excelize.ToAlphaString(icol) + strconv.Itoa(irow+1)
					if checkCellInArea(cellname, cellRef) {
						celldata = xlsx.GetCellValue(SheetName, ref[0])
						break
					}
				}
			}
			//check name column
			colname := xlsx.GetCellValue(SheetName, excelize.ToAlphaString(icol)+strconv.Itoa(headerrow))
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
	if filetype == "dn" {
		hangtondn = hdata
		errmsgdn = ""
		updatetimedn = time.Now()
	} else if filetype == "hn" {
		hangtonhn = hdata
		errmsghn = ""
		updatetimehn = time.Now()
	} else {
		hangtonsg = hdata
		errmsg = ""
		updatetime = time.Now()
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
