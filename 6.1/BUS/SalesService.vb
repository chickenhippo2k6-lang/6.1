Imports System.Linq
Imports DAL
Imports DTO

Namespace Global.BUS
    Public Class SalesService
        Private ReadOnly _repository As SalesRepository

        Public Sub New()
            _repository = New SalesRepository()
        End Sub

        Public Function GetProducts(searchText As String) As List(Of ProductDto)
            Dim query = If(searchText, String.Empty).Trim()
            If query.Length > 150 Then Throw New ArgumentException("Từ khóa tìm kiếm tối đa 150 ký tự.")
            Return _repository.GetProducts(query)
        End Function

        Public Sub SaveProduct(product As ProductDto, isNew As Boolean)
            If product Is Nothing Then Throw New ArgumentNullException(NameOf(product))
            product.MaSP = If(product.MaSP, String.Empty).Trim()
            product.TenSP = If(product.TenSP, String.Empty).Trim()
            If product.MaSP.Length = 0 OrElse product.MaSP.Length > 20 Then Throw New ArgumentException("Mã sản phẩm bắt buộc và tối đa 20 ký tự.")
            If product.TenSP.Length = 0 OrElse product.TenSP.Length > 150 Then Throw New ArgumentException("Tên sản phẩm bắt buộc và tối đa 150 ký tự.")
            If product.GiaBan <= 0D Then Throw New ArgumentException("Giá bán phải lớn hơn 0.")
            If Decimal.Round(product.GiaBan, 2) <> product.GiaBan Then Throw New ArgumentException("Giá bán chỉ được có tối đa 2 chữ số thập phân.")
            If product.SoLuongTon < 0 Then Throw New ArgumentException("Tồn kho không được âm.")
            If isNew Then
                _repository.AddProduct(product)
            Else
                _repository.UpdateProduct(product)
            End If
        End Sub

        Public Sub DeleteProduct(productId As String)
            If String.IsNullOrWhiteSpace(productId) Then Throw New ArgumentException("Chọn sản phẩm cần xóa.")
            _repository.DeleteProduct(productId.Trim())
        End Sub

        Public Function GetCustomers() As List(Of CustomerDto)
            Return _repository.GetCustomers()
        End Function

        Public Sub SaveCustomer(customer As CustomerDto, isNew As Boolean)
            If customer Is Nothing Then Throw New ArgumentNullException(NameOf(customer))
            customer.TenKH = If(customer.TenKH, String.Empty).Trim()
            customer.DienThoai = If(customer.DienThoai, String.Empty).Trim()
            If customer.TenKH.Length = 0 OrElse customer.TenKH.Length > 100 Then Throw New ArgumentException("Tên khách hàng bắt buộc và tối đa 100 ký tự.")
            If customer.DienThoai.Length > 15 Then Throw New ArgumentException("Số điện thoại tối đa 15 ký tự.")
            If isNew Then
                _repository.AddCustomer(customer)
            Else
                _repository.UpdateCustomer(customer)
            End If
        End Sub

        Public Sub DeleteCustomer(customerId As Integer)
            If customerId <= 0 Then Throw New ArgumentException("Chọn khách hàng cần xóa.")
            _repository.DeleteCustomer(customerId)
        End Sub

        Public Function GetEmployees() As List(Of EmployeeDto)
            Return _repository.GetEmployees()
        End Function

        Public Sub SaveEmployee(employee As EmployeeDto, isNew As Boolean)
            If employee Is Nothing Then Throw New ArgumentNullException(NameOf(employee))
            employee.TenNV = If(employee.TenNV, String.Empty).Trim()
            employee.SoDienThoai = If(employee.SoDienThoai, String.Empty).Trim()
            employee.ChucVu = If(employee.ChucVu, String.Empty).Trim()
            If employee.TenNV.Length = 0 OrElse employee.TenNV.Length > 100 Then Throw New ArgumentException("Tên nhân viên bắt buộc và tối đa 100 ký tự.")
            If employee.SoDienThoai.Length > 15 Then Throw New ArgumentException("Số điện thoại tối đa 15 ký tự.")
            If employee.ChucVu.Length > 50 Then Throw New ArgumentException("Chức vụ tối đa 50 ký tự.")
            If isNew Then
                _repository.AddEmployee(employee)
            Else
                _repository.UpdateEmployee(employee)
            End If
        End Sub

        Public Sub DeleteEmployee(employeeId As Integer)
            If employeeId <= 0 Then Throw New ArgumentException("Chọn nhân viên cần xóa.")
            _repository.DeleteEmployee(employeeId)
        End Sub

        Public Function Checkout(employeeId As Integer, customerId As Integer?, cart As IEnumerable(Of CartItemDto)) As CheckoutResultDto
            If employeeId <= 0 Then Throw New ArgumentException("Chọn nhân viên bán hàng.")
            If cart Is Nothing Then Throw New ArgumentNullException(NameOf(cart))

            Dim currentProducts = _repository.GetProducts(String.Empty).ToDictionary(Function(product) product.MaSP, StringComparer.OrdinalIgnoreCase)
            Dim normalizedCart As New List(Of CartItemDto)()
            For Each grouped In cart.GroupBy(Function(item) item.MaSP, StringComparer.OrdinalIgnoreCase)
                Dim quantity As Integer = 0
                For Each item In grouped
                    If item.SoLuong <= 0 Then Throw New ArgumentException("Số lượng bán phải lớn hơn 0.")
                    quantity = CInt(CLng(quantity) + item.SoLuong)
                Next
                Dim product As ProductDto = Nothing
                If Not currentProducts.TryGetValue(grouped.Key, product) Then Throw New InvalidOperationException($"Không tìm thấy sản phẩm '{grouped.Key}'.")
                normalizedCart.Add(New CartItemDto With {
                    .MaSP = product.MaSP,
                    .TenSP = product.TenSP,
                    .SoLuong = quantity,
                    .DonGia = product.GiaBan
                })
            Next
            If normalizedCart.Count = 0 Then Throw New ArgumentException("Giỏ hàng chưa có sản phẩm.")

            Return _repository.Checkout(employeeId, customerId, normalizedCart)
        End Function

        Public Function GetDailyRevenue(day As DateTime) As Decimal
            Return _repository.GetDailyRevenue(day)
        End Function

        Public Function GetMonthlyRevenue(month As DateTime) As Decimal
            Return _repository.GetMonthlyRevenue(month)
        End Function

        Public Function GetTopProducts() As List(Of TopProductDto)
            Return _repository.GetTopProducts()
        End Function
    End Class
End Namespace
