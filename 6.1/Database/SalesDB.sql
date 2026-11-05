IF DB_ID(N'SalesDB') IS NULL
BEGIN
	CREATE DATABASE SalesDB;
END;
GO

USE SalesDB;
GO

IF OBJECT_ID(N'dbo.NhanVien', N'U') IS NULL
BEGIN
	CREATE TABLE dbo.NhanVien
	(
		MaNV int IDENTITY(1,1) NOT NULL CONSTRAINT PK_NhanVien PRIMARY KEY,
		TenNV nvarchar(100) NOT NULL,
		SoDienThoai varchar(15) NULL,
		ChucVu nvarchar(50) NULL
	);
	CREATE UNIQUE INDEX UX_NhanVien_SoDienThoai ON dbo.NhanVien(SoDienThoai) WHERE SoDienThoai IS NOT NULL;
END;
GO

IF OBJECT_ID(N'dbo.KhachHang', N'U') IS NULL
BEGIN
	CREATE TABLE dbo.KhachHang
	(
		MaKH int IDENTITY(1,1) NOT NULL CONSTRAINT PK_KhachHang PRIMARY KEY,
		TenKH nvarchar(100) NOT NULL,
		DienThoai varchar(15) NULL,
		DiemTichLuy int NOT NULL CONSTRAINT DF_KhachHang_DiemTichLuy DEFAULT (0),
		CONSTRAINT CK_KhachHang_DiemTichLuy CHECK (DiemTichLuy >= 0)
	);
	CREATE UNIQUE INDEX UX_KhachHang_DienThoai ON dbo.KhachHang(DienThoai) WHERE DienThoai IS NOT NULL;
END;
GO

IF OBJECT_ID(N'dbo.SanPham', N'U') IS NULL
BEGIN
	CREATE TABLE dbo.SanPham
	(
		MaSP varchar(20) NOT NULL CONSTRAINT PK_SanPham PRIMARY KEY,
		TenSP nvarchar(150) NOT NULL,
		GiaBan decimal(18,2) NOT NULL,
		SoLuongTon int NOT NULL,
		CONSTRAINT CK_SanPham_GiaBan CHECK (GiaBan > 0),
		CONSTRAINT CK_SanPham_SoLuongTon CHECK (SoLuongTon >= 0)
	);
END;
GO

IF OBJECT_ID(N'dbo.HoaDon', N'U') IS NULL
BEGIN
	CREATE TABLE dbo.HoaDon
	(
		MaHD int IDENTITY(1,1) NOT NULL CONSTRAINT PK_HoaDon PRIMARY KEY,
		NgayLap datetime NOT NULL CONSTRAINT DF_HoaDon_NgayLap DEFAULT (GETDATE()),
		MaNV int NOT NULL,
		MaKH int NULL,
		TongTien decimal(18,2) NOT NULL,
		CONSTRAINT FK_HoaDon_NhanVien FOREIGN KEY (MaNV) REFERENCES dbo.NhanVien(MaNV),
		CONSTRAINT FK_HoaDon_KhachHang FOREIGN KEY (MaKH) REFERENCES dbo.KhachHang(MaKH) ON DELETE SET NULL,
		CONSTRAINT CK_HoaDon_TongTien CHECK (TongTien >= 0)
	);
END;
GO

IF OBJECT_ID(N'dbo.ChiTietHoaDon', N'U') IS NULL
BEGIN
	CREATE TABLE dbo.ChiTietHoaDon
	(
		MaHD int NOT NULL,
		MaSP varchar(20) NOT NULL,
		SoLuong int NOT NULL,
		DonGia decimal(18,2) NOT NULL,
		CONSTRAINT PK_ChiTietHoaDon PRIMARY KEY (MaHD, MaSP),
		CONSTRAINT FK_ChiTietHoaDon_HoaDon FOREIGN KEY (MaHD) REFERENCES dbo.HoaDon(MaHD),
		CONSTRAINT FK_ChiTietHoaDon_SanPham FOREIGN KEY (MaSP) REFERENCES dbo.SanPham(MaSP),
		CONSTRAINT CK_ChiTietHoaDon_SoLuong CHECK (SoLuong > 0),
		CONSTRAINT CK_ChiTietHoaDon_DonGia CHECK (DonGia > 0)
	);
END;
GO
