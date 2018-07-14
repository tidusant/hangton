package main

type HangTonData struct {
	MaNhomHang1       string
	MaNhomHang2       string
	MaNhomHang3       string
	MaNhomHang4       string
	Kho               string
	MaHang            string
	TenHang           string
	Dvt               string
	TonCuoiSL         int
	Tong2Kho          int
	GiaHoreca         int
	BanTBThang        int
	UocLuongBan4Thang int
	SLCanDauKy        int
	SLCanHienTai      int
	TL1               int
	TL2               int
	TL3               int
	TL4               int
	TL5               int
	TL6               int
}

type HangTonDataReturn struct {
	MaHang  string
	TenHang string

	Tong2Kho int

	UocLuongBan4Thang int

	TL1 int
	TL2 int
	TL3 int
	TL4 int
	TL5 int
	TL6 int
}
