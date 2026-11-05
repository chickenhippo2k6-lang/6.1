Namespace Global.DTO
    Public Class ProductDto
        Public Property MaSP As String = String.Empty
        Public Property TenSP As String = String.Empty
        Public Property GiaBan As Decimal
        Public Property SoLuongTon As Integer
    End Class

    Public Class CustomerDto
        Public Property MaKH As Integer
        Public Property TenKH As String = String.Empty
        Public Property DienThoai As String = String.Empty
        Public Property DiemTichLuy As Integer
    End Class

    Public Class EmployeeDto
        Public Property MaNV As Integer
        Public Property TenNV As String = String.Empty
        Public Property SoDienThoai As String = String.Empty
        Public Property ChucVu As String = String.Empty
    End Class

    Public Class CartItemDto
        Public Property MaSP As String = String.Empty
        Public Property TenSP As String = String.Empty
        Public Property SoLuong As Integer
        Public Property DonGia As Decimal
        Public ReadOnly Property ThanhTien As Decimal
            Get
                Return SoLuong * DonGia
            End Get
        End Property
    End Class

    Public Class TopProductDto
        Public Property MaSP As String = String.Empty
        Public Property TenSP As String = String.Empty
        Public Property SoLuongBan As Integer
    End Class

    Public Class CheckoutResultDto
        Public Property MaHD As Integer
        Public Property TongTien As Decimal
        Public Property DiemCong As Integer
    End Class
End Namespace
