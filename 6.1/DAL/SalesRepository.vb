Imports System.Data
Imports System.Linq
Imports DTO
Imports Microsoft.Data.SqlClient

Namespace Global.DAL
    Public Class SalesRepository
        Private ReadOnly _connectionString As String

        Public Sub New()
            _connectionString = Environment.GetEnvironmentVariable("SALESDB_CONNECTION_STRING")
            If String.IsNullOrWhiteSpace(_connectionString) Then
                _connectionString = "Server=(localdb)\MSSQLLocalDB;Database=SalesDB;Integrated Security=True;Encrypt=True;TrustServerCertificate=True"
            End If
        End Sub

        Public Function GetProducts(searchText As String) As List(Of ProductDto)
            Dim results As New List(Of ProductDto)()
            Using connection As New SqlConnection(_connectionString)
                Using command As New SqlCommand("SELECT MaSP, TenSP, GiaBan, SoLuongTon FROM dbo.SanPham WHERE (@search = N'' OR TenSP LIKE @pattern OR MaSP LIKE @pattern) ORDER BY TenSP", connection)
                    AddParameter(command, "@search", SqlDbType.NVarChar, If(searchText, String.Empty), 150)
                    AddParameter(command, "@pattern", SqlDbType.NVarChar, "%" & If(searchText, String.Empty) & "%", 152)
                    connection.Open()
                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            results.Add(New ProductDto With {
                                .MaSP = reader.GetString(0),
                                .TenSP = reader.GetString(1),
                                .GiaBan = reader.GetDecimal(2),
                                .SoLuongTon = reader.GetInt32(3)
                            })
                        End While
                    End Using
                End Using
            End Using
            Return results
        End Function

        Public Sub AddProduct(product As ProductDto)
            Using connection As New SqlConnection(_connectionString)
                Using command As New SqlCommand("INSERT INTO dbo.SanPham (MaSP, TenSP, GiaBan, SoLuongTon) VALUES (@id, @name, @price, @stock)", connection)
                    AddProductParameters(command, product)
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub UpdateProduct(product As ProductDto)
            Using connection As New SqlConnection(_connectionString)
                Using command As New SqlCommand("UPDATE dbo.SanPham SET TenSP = @name, GiaBan = @price, SoLuongTon = @stock WHERE MaSP = @id", connection)
                    AddProductParameters(command, product)
                    connection.Open()
                    If command.ExecuteNonQuery() = 0 Then Throw New InvalidOperationException("Không tìm thấy sản phẩm cần cập nhật.")
                End Using
            End Using
        End Sub

        Public Sub DeleteProduct(productId As String)
            ExecuteById("DELETE FROM dbo.SanPham WHERE MaSP = @id", "@id", SqlDbType.VarChar, productId, 20)
        End Sub

        Public Function GetCustomers() As List(Of CustomerDto)
            Dim results As New List(Of CustomerDto)()
            Using connection As New SqlConnection(_connectionString)
                Using command As New SqlCommand("SELECT MaKH, TenKH, DienThoai, DiemTichLuy FROM dbo.KhachHang ORDER BY TenKH", connection)
                    connection.Open()
                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            results.Add(New CustomerDto With {
                                .MaKH = reader.GetInt32(0),
                                .TenKH = reader.GetString(1),
                                .DienThoai = If(reader.IsDBNull(2), String.Empty, reader.GetString(2)),
                                .DiemTichLuy = reader.GetInt32(3)
                            })
                        End While
                    End Using
                End Using
            End Using
            Return results
        End Function

        Public Sub AddCustomer(customer As CustomerDto)
            Using connection As New SqlConnection(_connectionString)
                Using command As New SqlCommand("INSERT INTO dbo.KhachHang (TenKH, DienThoai) VALUES (@name, @phone)", connection)
                    AddParameter(command, "@name", SqlDbType.NVarChar, customer.TenKH, 100)
                    AddParameter(command, "@phone", SqlDbType.VarChar, OptionalValue(customer.DienThoai), 15)
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub UpdateCustomer(customer As CustomerDto)
            Using connection As New SqlConnection(_connectionString)
                Using command As New SqlCommand("UPDATE dbo.KhachHang SET TenKH = @name, DienThoai = @phone WHERE MaKH = @id", connection)
                    AddParameter(command, "@name", SqlDbType.NVarChar, customer.TenKH, 100)
                    AddParameter(command, "@phone", SqlDbType.VarChar, OptionalValue(customer.DienThoai), 15)
                    AddParameter(command, "@id", SqlDbType.Int, customer.MaKH)
                    connection.Open()
                    If command.ExecuteNonQuery() = 0 Then Throw New InvalidOperationException("Không tìm thấy khách hàng cần cập nhật.")
                End Using
            End Using
        End Sub

        Public Sub DeleteCustomer(customerId As Integer)
            ExecuteById("DELETE FROM dbo.KhachHang WHERE MaKH = @id", "@id", SqlDbType.Int, customerId)
        End Sub

        Public Function GetEmployees() As List(Of EmployeeDto)
            Dim results As New List(Of EmployeeDto)()
            Using connection As New SqlConnection(_connectionString)
                Using command As New SqlCommand("SELECT MaNV, TenNV, SoDienThoai, ChucVu FROM dbo.NhanVien ORDER BY TenNV", connection)
                    connection.Open()
                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            results.Add(New EmployeeDto With {
                                .MaNV = reader.GetInt32(0),
                                .TenNV = reader.GetString(1),
                                .SoDienThoai = If(reader.IsDBNull(2), String.Empty, reader.GetString(2)),
                                .ChucVu = If(reader.IsDBNull(3), String.Empty, reader.GetString(3))
                            })
                        End While
                    End Using
                End Using
            End Using
            Return results
        End Function

        Public Sub AddEmployee(employee As EmployeeDto)
            Using connection As New SqlConnection(_connectionString)
                Using command As New SqlCommand("INSERT INTO dbo.NhanVien (TenNV, SoDienThoai, ChucVu) VALUES (@name, @phone, @role)", connection)
                    AddEmployeeParameters(command, employee, False)
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub UpdateEmployee(employee As EmployeeDto)
            Using connection As New SqlConnection(_connectionString)
                Using command As New SqlCommand("UPDATE dbo.NhanVien SET TenNV = @name, SoDienThoai = @phone, ChucVu = @role WHERE MaNV = @id", connection)
                    AddEmployeeParameters(command, employee, True)
                    connection.Open()
                    If command.ExecuteNonQuery() = 0 Then Throw New InvalidOperationException("Không tìm thấy nhân viên cần cập nhật.")
                End Using
            End Using
        End Sub

        Public Sub DeleteEmployee(employeeId As Integer)
            ExecuteById("DELETE FROM dbo.NhanVien WHERE MaNV = @id", "@id", SqlDbType.Int, employeeId)
        End Sub

        Public Function GetDailyRevenue(day As DateTime) As Decimal
            Return GetRevenue(day.Date, day.Date.AddDays(1))
        End Function

        Public Function GetMonthlyRevenue(month As DateTime) As Decimal
            Dim firstDay = New DateTime(month.Year, month.Month, 1)
            Return GetRevenue(firstDay, firstDay.AddMonths(1))
        End Function

        Public Function GetTopProducts() As List(Of TopProductDto)
            Dim results As New List(Of TopProductDto)()
            Using connection As New SqlConnection(_connectionString)
                Using command As New SqlCommand("SELECT TOP (5) sp.MaSP, sp.TenSP, SUM(ct.SoLuong) AS SoLuongBan FROM dbo.ChiTietHoaDon AS ct INNER JOIN dbo.SanPham AS sp ON sp.MaSP = ct.MaSP GROUP BY sp.MaSP, sp.TenSP ORDER BY SUM(ct.SoLuong) DESC, sp.TenSP", connection)
                    connection.Open()
                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            results.Add(New TopProductDto With {
                                .MaSP = reader.GetString(0),
                                .TenSP = reader.GetString(1),
                                .SoLuongBan = reader.GetInt32(2)
                            })
                        End While
                    End Using
                End Using
            End Using
            Return results
        End Function

        Public Function Checkout(employeeId As Integer, customerId As Integer?, items As IReadOnlyCollection(Of CartItemDto)) As CheckoutResultDto
            Dim result As CheckoutResultDto = Nothing
            Using connection As New SqlConnection(_connectionString)
                connection.Open()
                Using transaction = connection.BeginTransaction(IsolationLevel.Serializable)
                    Try
                        Dim saleItems As New List(Of CartItemDto)()
                        For Each item In items
                            Using productCommand As New SqlCommand("SELECT GiaBan, SoLuongTon FROM dbo.SanPham WITH (UPDLOCK, HOLDLOCK) WHERE MaSP = @productId", connection, transaction)
                                AddParameter(productCommand, "@productId", SqlDbType.VarChar, item.MaSP, 20)
                                Using reader = productCommand.ExecuteReader()
                                    If Not reader.Read() OrElse reader.GetInt32(1) < item.SoLuong Then
                                        Throw New InvalidOperationException($"Sản phẩm '{item.MaSP}' không tồn tại hoặc không đủ tồn kho.")
                                    End If
                                    saleItems.Add(New CartItemDto With {
                                        .MaSP = item.MaSP,
                                        .TenSP = item.TenSP,
                                        .SoLuong = item.SoLuong,
                                        .DonGia = reader.GetDecimal(0)
                                    })
                                End Using
                            End Using
                        Next

                        Dim total As Decimal = saleItems.Sum(Function(item) item.ThanhTien)
                        Dim points = CInt(Math.Floor(total / 100000D))
                        result = New CheckoutResultDto With {.TongTien = total, .DiemCong = points}

                        Using invoiceCommand As New SqlCommand("INSERT INTO dbo.HoaDon (MaNV, MaKH, TongTien) OUTPUT INSERTED.MaHD VALUES (@employeeId, @customerId, @total)", connection, transaction)
                            AddParameter(invoiceCommand, "@employeeId", SqlDbType.Int, employeeId)
                            AddParameter(invoiceCommand, "@customerId", SqlDbType.Int, If(customerId.HasValue, CType(customerId.Value, Object), DBNull.Value))
                            AddDecimalParameter(invoiceCommand, "@total", total)
                            result.MaHD = CInt(invoiceCommand.ExecuteScalar())
                        End Using

                        For Each item In saleItems
                            Using detailCommand As New SqlCommand("INSERT INTO dbo.ChiTietHoaDon (MaHD, MaSP, SoLuong, DonGia) VALUES (@invoiceId, @productId, @quantity, @price)", connection, transaction)
                                AddParameter(detailCommand, "@invoiceId", SqlDbType.Int, result.MaHD)
                                AddParameter(detailCommand, "@productId", SqlDbType.VarChar, item.MaSP, 20)
                                AddParameter(detailCommand, "@quantity", SqlDbType.Int, item.SoLuong)
                                AddDecimalParameter(detailCommand, "@price", item.DonGia)
                                detailCommand.ExecuteNonQuery()
                            End Using

                            Using stockCommand As New SqlCommand("UPDATE dbo.SanPham SET SoLuongTon = SoLuongTon - @quantity WHERE MaSP = @productId AND SoLuongTon >= @quantity", connection, transaction)
                                AddParameter(stockCommand, "@quantity", SqlDbType.Int, item.SoLuong)
                                AddParameter(stockCommand, "@productId", SqlDbType.VarChar, item.MaSP, 20)
                                If stockCommand.ExecuteNonQuery() <> 1 Then
                                    Throw New InvalidOperationException($"Sản phẩm '{item.MaSP}' không tồn tại hoặc không đủ tồn kho.")
                                End If
                            End Using
                        Next

                        If customerId.HasValue AndAlso points > 0 Then
                            Using pointsCommand As New SqlCommand("UPDATE dbo.KhachHang SET DiemTichLuy = DiemTichLuy + @points WHERE MaKH = @customerId", connection, transaction)
                                AddParameter(pointsCommand, "@points", SqlDbType.Int, points)
                                AddParameter(pointsCommand, "@customerId", SqlDbType.Int, customerId.Value)
                                If pointsCommand.ExecuteNonQuery() <> 1 Then Throw New InvalidOperationException("Không tìm thấy khách hàng.")
                            End Using
                        End If

                        transaction.Commit()
                    Catch
                        Try
                            transaction.Rollback()
                        Catch
                        End Try
                        Throw
                    End Try
                End Using
            End Using
            Return result
        End Function

        Private Function GetRevenue(startDate As DateTime, endDate As DateTime) As Decimal
            Using connection As New SqlConnection(_connectionString)
                Using command As New SqlCommand("SELECT COALESCE(SUM(TongTien), 0) FROM dbo.HoaDon WHERE NgayLap >= @startDate AND NgayLap < @endDate", connection)
                    AddParameter(command, "@startDate", SqlDbType.DateTime, startDate)
                    AddParameter(command, "@endDate", SqlDbType.DateTime, endDate)
                    connection.Open()
                    Return Convert.ToDecimal(command.ExecuteScalar())
                End Using
            End Using
        End Function

        Private Shared Sub AddProductParameters(command As SqlCommand, product As ProductDto)
            AddParameter(command, "@id", SqlDbType.VarChar, product.MaSP, 20)
            AddParameter(command, "@name", SqlDbType.NVarChar, product.TenSP, 150)
            AddDecimalParameter(command, "@price", product.GiaBan)
            AddParameter(command, "@stock", SqlDbType.Int, product.SoLuongTon)
        End Sub

        Private Shared Sub AddEmployeeParameters(command As SqlCommand, employee As EmployeeDto, includeId As Boolean)
            AddParameter(command, "@name", SqlDbType.NVarChar, employee.TenNV, 100)
            AddParameter(command, "@phone", SqlDbType.VarChar, OptionalValue(employee.SoDienThoai), 15)
            AddParameter(command, "@role", SqlDbType.NVarChar, OptionalValue(employee.ChucVu), 50)
            If includeId Then AddParameter(command, "@id", SqlDbType.Int, employee.MaNV)
        End Sub

        Private Shared Sub AddParameter(command As SqlCommand, name As String, type As SqlDbType, value As Object, Optional size As Integer = 0)
            Dim parameter = command.Parameters.Add(name, type)
            If size > 0 Then parameter.Size = size
            parameter.Value = If(value Is Nothing, DBNull.Value, value)
        End Sub

        Private Shared Sub AddDecimalParameter(command As SqlCommand, name As String, value As Decimal)
            Dim parameter = command.Parameters.Add(name, SqlDbType.Decimal)
            parameter.Precision = 18
            parameter.Scale = 2
            parameter.Value = value
        End Sub

        Private Sub ExecuteById(sql As String, parameterName As String, type As SqlDbType, value As Object, Optional size As Integer = 0)
            Using connection As New SqlConnection(_connectionString)
                Using command As New SqlCommand(sql, connection)
                    AddParameter(command, parameterName, type, value, size)
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Private Shared Function OptionalValue(value As String) As Object
            If String.IsNullOrWhiteSpace(value) Then Return DBNull.Value
            Return value.Trim()
        End Function
    End Class
End Namespace
