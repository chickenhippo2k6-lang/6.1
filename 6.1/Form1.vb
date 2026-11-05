Imports System.Linq
Imports BUS
Imports DTO

Public Class Form1
    Private ReadOnly _service As New SalesService()
    Private ReadOnly _cart As New List(Of CartItemDto)()
    Private _products As List(Of ProductDto) = New List(Of ProductDto)()
    Private _selectedCustomerId As Integer
    Private _selectedEmployeeId As Integer

    Private ReadOnly _productGrid As DataGridView = CreateGrid()
    Private ReadOnly _customerGrid As DataGridView = CreateGrid()
    Private ReadOnly _employeeGrid As DataGridView = CreateGrid()
    Private ReadOnly _cartGrid As DataGridView = CreateGrid()
    Private ReadOnly _topProductsGrid As DataGridView = CreateGrid()

    Private ReadOnly _productSearch As New TextBox()
    Private ReadOnly _productId As New TextBox()
    Private ReadOnly _productName As New TextBox()
    Private ReadOnly _productPrice As New NumericUpDown()
    Private ReadOnly _productStock As New NumericUpDown()

    Private ReadOnly _customerName As New TextBox()
    Private ReadOnly _customerPhone As New TextBox()
    Private ReadOnly _employeeName As New TextBox()
    Private ReadOnly _employeePhone As New TextBox()
    Private ReadOnly _employeeRole As New TextBox()

    Private ReadOnly _employeeSelect As New ComboBox()
    Private ReadOnly _customerSelect As New ComboBox()
    Private ReadOnly _saleProductCode As New TextBox()
    Private ReadOnly _saleQuantity As New NumericUpDown()
    Private ReadOnly _dailyDate As New DateTimePicker()
    Private ReadOnly _monthlyDate As New DateTimePicker()
    Private ReadOnly _reportResult As New Label()

    Public Sub New()
        InitializeComponent()
        Text = "Hệ thống quản lý bán hàng - Siêu thị mini"
        StartPosition = FormStartPosition.CenterScreen
        MinimumSize = New Size(900, 600)
        BuildInterface()
        AddHandler Shown, AddressOf FormShown
    End Sub

    Private Sub BuildInterface()
        Dim tabs As New TabControl With {.Dock = DockStyle.Fill}
        tabs.TabPages.Add(BuildProductsTab())
        tabs.TabPages.Add(BuildCustomersTab())
        tabs.TabPages.Add(BuildEmployeesTab())
        tabs.TabPages.Add(BuildSalesTab())
        tabs.TabPages.Add(BuildReportsTab())
        Controls.Add(tabs)

        ConfigureNumeric(_productPrice, 9999999999999999D, 2)
        ConfigureNumeric(_productStock, 2147483647D, 0)
        ConfigureNumeric(_saleQuantity, 2147483647D, 0)
        _saleQuantity.Minimum = 1
    End Sub

    Private Function BuildProductsTab() As TabPage
        Dim page As New TabPage("Sản phẩm")
        Dim searchPanel As New FlowLayoutPanel With {.Dock = DockStyle.Top, .Height = 42, .Padding = New Padding(6), .WrapContents = False}
        _productSearch.Width = 260
        _productSearch.MaxLength = 150
        searchPanel.Controls.Add(New Label With {.Text = "Tên hoặc mã sản phẩm:", .AutoSize = True, .Margin = New Padding(3, 8, 3, 3)})
        searchPanel.Controls.Add(_productSearch)
        searchPanel.Controls.Add(NewButton("Tìm kiếm", Sub() RunUiAction(Sub() RefreshProducts(_productSearch.Text))))
        searchPanel.Controls.Add(NewButton("Làm mới", Sub() RunUiAction(AddressOf RefreshProductsAll)))

        Dim editor As New FlowLayoutPanel With {.Dock = DockStyle.Bottom, .Height = 112, .Padding = New Padding(6), .WrapContents = True}
        _productId.Width = 110
        _productId.MaxLength = 20
        _productName.Width = 180
        _productName.MaxLength = 150
        _productPrice.Width = 145
        _productStock.Width = 110
        AddLabeledControl(editor, "Mã SP", _productId)
        AddLabeledControl(editor, "Tên SP", _productName)
        AddLabeledControl(editor, "Giá bán", _productPrice)
        AddLabeledControl(editor, "Tồn kho", _productStock)
        editor.Controls.Add(NewButton("Thêm", Sub() RunUiAction(AddressOf AddProduct)))
        editor.Controls.Add(NewButton("Cập nhật", Sub() RunUiAction(AddressOf UpdateProduct)))
        editor.Controls.Add(NewButton("Xóa", Sub() RunUiAction(AddressOf DeleteProduct)))
        editor.Controls.Add(NewButton("Xóa nhập liệu", Sub() ClearProductFields()))

        _productGrid.Dock = DockStyle.Fill
        AddHandler _productGrid.SelectionChanged, AddressOf ProductSelectionChanged
        page.Controls.Add(_productGrid)
        page.Controls.Add(editor)
        page.Controls.Add(searchPanel)
        Return page
    End Function

    Private Function BuildCustomersTab() As TabPage
        Dim page As New TabPage("Khách hàng")
        Dim editor As New FlowLayoutPanel With {.Dock = DockStyle.Bottom, .Height = 100, .Padding = New Padding(6), .WrapContents = True}
        _customerName.Width = 220
        _customerName.MaxLength = 100
        _customerPhone.Width = 150
        _customerPhone.MaxLength = 15
        AddLabeledControl(editor, "Tên KH", _customerName)
        AddLabeledControl(editor, "Điện thoại", _customerPhone)
        editor.Controls.Add(NewButton("Thêm", Sub() RunUiAction(AddressOf AddCustomer)))
        editor.Controls.Add(NewButton("Cập nhật", Sub() RunUiAction(AddressOf UpdateCustomer)))
        editor.Controls.Add(NewButton("Xóa", Sub() RunUiAction(AddressOf DeleteCustomer)))
        editor.Controls.Add(NewButton("Xóa nhập liệu", Sub() ClearCustomerFields()))
        _customerGrid.Dock = DockStyle.Fill
        AddHandler _customerGrid.SelectionChanged, AddressOf CustomerSelectionChanged
        page.Controls.Add(_customerGrid)
        page.Controls.Add(editor)
        Return page
    End Function

    Private Function BuildEmployeesTab() As TabPage
        Dim page As New TabPage("Nhân viên")
        Dim editor As New FlowLayoutPanel With {.Dock = DockStyle.Bottom, .Height = 100, .Padding = New Padding(6), .WrapContents = True}
        _employeeName.Width = 200
        _employeeName.MaxLength = 100
        _employeePhone.Width = 140
        _employeePhone.MaxLength = 15
        _employeeRole.Width = 140
        _employeeRole.MaxLength = 50
        AddLabeledControl(editor, "Tên NV", _employeeName)
        AddLabeledControl(editor, "Điện thoại", _employeePhone)
        AddLabeledControl(editor, "Chức vụ", _employeeRole)
        editor.Controls.Add(NewButton("Thêm", Sub() RunUiAction(AddressOf AddEmployee)))
        editor.Controls.Add(NewButton("Cập nhật", Sub() RunUiAction(AddressOf UpdateEmployee)))
        editor.Controls.Add(NewButton("Xóa", Sub() RunUiAction(AddressOf DeleteEmployee)))
        editor.Controls.Add(NewButton("Xóa nhập liệu", Sub() ClearEmployeeFields()))
        _employeeGrid.Dock = DockStyle.Fill
        AddHandler _employeeGrid.SelectionChanged, AddressOf EmployeeSelectionChanged
        page.Controls.Add(_employeeGrid)
        page.Controls.Add(editor)
        Return page
    End Function

    Private Function BuildSalesTab() As TabPage
        Dim page As New TabPage("Bán hàng")
        Dim topPanel As New FlowLayoutPanel With {.Dock = DockStyle.Top, .Height = 82, .Padding = New Padding(6), .WrapContents = True}
        _employeeSelect.DropDownStyle = ComboBoxStyle.DropDownList
        _employeeSelect.Width = 190
        _customerSelect.DropDownStyle = ComboBoxStyle.DropDownList
        _customerSelect.Width = 220
        _saleProductCode.Width = 140
        _saleQuantity.Width = 90
        AddLabeledControl(topPanel, "Nhân viên", _employeeSelect)
        AddLabeledControl(topPanel, "Khách hàng", _customerSelect)
        AddLabeledControl(topPanel, "Mã sản phẩm", _saleProductCode)
        AddLabeledControl(topPanel, "Số lượng", _saleQuantity)
        topPanel.Controls.Add(NewButton("Thêm vào giỏ", Sub() RunUiAction(AddressOf AddCartItem)))
        topPanel.Controls.Add(NewButton("Xóa dòng chọn", Sub() RemoveCartItem()))

        Dim bottomPanel As New FlowLayoutPanel With {.Dock = DockStyle.Bottom, .Height = 55, .Padding = New Padding(8)}
        bottomPanel.Controls.Add(NewButton("Thanh toán", Sub() RunUiAction(AddressOf Checkout)))
        bottomPanel.Controls.Add(NewButton("Xóa giỏ hàng", Sub() ClearCart()))
        _cartGrid.Dock = DockStyle.Fill
        page.Controls.Add(_cartGrid)
        page.Controls.Add(bottomPanel)
        page.Controls.Add(topPanel)
        Return page
    End Function

    Private Function BuildReportsTab() As TabPage
        Dim page As New TabPage("Báo cáo")
        Dim toolbar As New FlowLayoutPanel With {.Dock = DockStyle.Top, .Height = 82, .Padding = New Padding(6), .WrapContents = True}
        _dailyDate.Format = DateTimePickerFormat.Custom
        _dailyDate.CustomFormat = "yyyy-MM-dd"
        _monthlyDate.Format = DateTimePickerFormat.Custom
        _monthlyDate.CustomFormat = "yyyy-MM"
        toolbar.Controls.Add(New Label With {.Text = "Ngày:", .AutoSize = True, .Margin = New Padding(3, 8, 3, 3)})
        toolbar.Controls.Add(_dailyDate)
        toolbar.Controls.Add(NewButton("Doanh thu ngày", Sub() RunUiAction(AddressOf ShowDailyRevenue)))
        toolbar.Controls.Add(New Label With {.Text = "Tháng:", .AutoSize = True, .Margin = New Padding(12, 8, 3, 3)})
        toolbar.Controls.Add(_monthlyDate)
        toolbar.Controls.Add(NewButton("Doanh thu tháng", Sub() RunUiAction(AddressOf ShowMonthlyRevenue)))
        toolbar.Controls.Add(NewButton("Top 5 bán chạy", Sub() RunUiAction(AddressOf RefreshTopProducts)))
        _reportResult.AutoSize = True
        _reportResult.Padding = New Padding(8)
        Dim resultPanel As New Panel With {.Dock = DockStyle.Bottom, .Height = 40}
        resultPanel.Controls.Add(_reportResult)
        _topProductsGrid.Dock = DockStyle.Fill
        page.Controls.Add(_topProductsGrid)
        page.Controls.Add(resultPanel)
        page.Controls.Add(toolbar)
        Return page
    End Function

    Private Sub FormShown(sender As Object, e As EventArgs)
        RunUiAction(AddressOf RefreshAll)
    End Sub

    Private Sub RefreshAll()
        RefreshProductsAll()
        RefreshCustomers()
        RefreshEmployees()
        RefreshSalesLookups()
        RefreshCart()
        RefreshTopProducts()
    End Sub

    Private Sub RefreshProductsAll()
        RefreshProducts(String.Empty)
    End Sub

    Private Sub RefreshProducts(searchText As String)
        _products = _service.GetProducts(searchText)
        _productGrid.DataSource = Nothing
        _productGrid.DataSource = _products.ToList()
        RefreshSaleProducts()
    End Sub

    Private Sub RefreshCustomers()
        _customerGrid.DataSource = Nothing
        _customerGrid.DataSource = _service.GetCustomers()
    End Sub

    Private Sub RefreshEmployees()
        _employeeGrid.DataSource = Nothing
        _employeeGrid.DataSource = _service.GetEmployees()
    End Sub

    Private Sub RefreshSalesLookups()
        Dim customers = _service.GetCustomers()
        customers.Insert(0, New CustomerDto With {.MaKH = 0, .TenKH = "Khách lẻ"})
        _customerSelect.DataSource = customers
        _customerSelect.DisplayMember = NameOf(CustomerDto.TenKH)
        _customerSelect.ValueMember = NameOf(CustomerDto.MaKH)

        _employeeSelect.DataSource = _service.GetEmployees()
        _employeeSelect.DisplayMember = NameOf(EmployeeDto.TenNV)
        _employeeSelect.ValueMember = NameOf(EmployeeDto.MaNV)
    End Sub

    Private Sub RefreshSaleProducts()
        Dim selectedCode = _saleProductCode.Text
        _saleProductCode.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        _saleProductCode.AutoCompleteSource = AutoCompleteSource.CustomSource
        Dim source As New AutoCompleteStringCollection()
        source.AddRange(_products.Select(Function(product) product.MaSP).ToArray())
        _saleProductCode.AutoCompleteCustomSource = source
        If selectedCode.Length > 0 Then _saleProductCode.Text = selectedCode
    End Sub

    Private Sub RefreshCart()
        _cartGrid.DataSource = Nothing
        _cartGrid.DataSource = _cart.Select(Function(item) New With {
            .MaSP = item.MaSP,
            .TenSP = item.TenSP,
            .SoLuong = item.SoLuong,
            .DonGia = item.DonGia,
            .ThanhTien = item.ThanhTien
        }).ToList()
    End Sub

    Private Sub RefreshTopProducts()
        _topProductsGrid.DataSource = Nothing
        _topProductsGrid.DataSource = _service.GetTopProducts()
    End Sub

    Private Sub AddProduct()
        _service.SaveProduct(ReadProduct(), True)
        RefreshProductsAll()
        ClearProductFields()
    End Sub

    Private Sub UpdateProduct()
        If String.IsNullOrWhiteSpace(_productId.Text) Then Throw New InvalidOperationException("Chọn hoặc nhập mã sản phẩm cần cập nhật.")
        _service.SaveProduct(ReadProduct(), False)
        RefreshProductsAll()
        ClearProductFields()
    End Sub

    Private Sub DeleteProduct()
        If String.IsNullOrWhiteSpace(_productId.Text) Then Throw New InvalidOperationException("Chọn sản phẩm cần xóa.")
        If MessageBox.Show("Xóa sản phẩm này?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        _service.DeleteProduct(_productId.Text)
        RefreshProductsAll()
        ClearProductFields()
    End Sub

    Private Function ReadProduct() As ProductDto
        Return New ProductDto With {
            .MaSP = _productId.Text,
            .TenSP = _productName.Text,
            .GiaBan = _productPrice.Value,
            .SoLuongTon = CInt(_productStock.Value)
        }
    End Function

    Private Sub AddCustomer()
        _service.SaveCustomer(New CustomerDto With {.TenKH = _customerName.Text, .DienThoai = _customerPhone.Text}, True)
        RefreshCustomers()
        RefreshSalesLookups()
        ClearCustomerFields()
    End Sub

    Private Sub UpdateCustomer()
        If _selectedCustomerId <= 0 Then Throw New InvalidOperationException("Chọn khách hàng cần cập nhật.")
        _service.SaveCustomer(New CustomerDto With {.MaKH = _selectedCustomerId, .TenKH = _customerName.Text, .DienThoai = _customerPhone.Text}, False)
        RefreshCustomers()
        RefreshSalesLookups()
        ClearCustomerFields()
    End Sub

    Private Sub DeleteCustomer()
        If _selectedCustomerId <= 0 Then Throw New InvalidOperationException("Chọn khách hàng cần xóa.")
        If MessageBox.Show("Xóa khách hàng này? Hóa đơn cũ sẽ chuyển sang khách lẻ.", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        _service.DeleteCustomer(_selectedCustomerId)
        RefreshCustomers()
        RefreshSalesLookups()
        ClearCustomerFields()
    End Sub

    Private Sub AddEmployee()
        _service.SaveEmployee(New EmployeeDto With {.TenNV = _employeeName.Text, .SoDienThoai = _employeePhone.Text, .ChucVu = _employeeRole.Text}, True)
        RefreshEmployees()
        RefreshSalesLookups()
        ClearEmployeeFields()
    End Sub

    Private Sub UpdateEmployee()
        If _selectedEmployeeId <= 0 Then Throw New InvalidOperationException("Chọn nhân viên cần cập nhật.")
        _service.SaveEmployee(New EmployeeDto With {.MaNV = _selectedEmployeeId, .TenNV = _employeeName.Text, .SoDienThoai = _employeePhone.Text, .ChucVu = _employeeRole.Text}, False)
        RefreshEmployees()
        RefreshSalesLookups()
        ClearEmployeeFields()
    End Sub

    Private Sub DeleteEmployee()
        If _selectedEmployeeId <= 0 Then Throw New InvalidOperationException("Chọn nhân viên cần xóa.")
        If MessageBox.Show("Xóa nhân viên này?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        _service.DeleteEmployee(_selectedEmployeeId)
        RefreshEmployees()
        RefreshSalesLookups()
        ClearEmployeeFields()
    End Sub

    Private Sub AddCartItem()
        Dim code = _saleProductCode.Text.Trim()
        Dim product = _products.FirstOrDefault(Function(candidate) String.Equals(candidate.MaSP, code, StringComparison.OrdinalIgnoreCase))
        If product Is Nothing Then Throw New InvalidOperationException("Không tìm thấy mã sản phẩm. Hãy làm mới danh sách sản phẩm.")
        Dim existing = _cart.FirstOrDefault(Function(item) String.Equals(item.MaSP, product.MaSP, StringComparison.OrdinalIgnoreCase))
        If existing Is Nothing Then
            _cart.Add(New CartItemDto With {.MaSP = product.MaSP, .TenSP = product.TenSP, .SoLuong = CInt(_saleQuantity.Value), .DonGia = product.GiaBan})
        Else
            existing.SoLuong = CInt(CLng(existing.SoLuong) + CInt(_saleQuantity.Value))
        End If
        RefreshCart()
        _saleProductCode.Clear()
    End Sub

    Private Sub RemoveCartItem()
        If _cartGrid.CurrentRow Is Nothing Then Return
        Dim code = Convert.ToString(_cartGrid.CurrentRow.Cells(NameOf(CartItemDto.MaSP)).Value)
        _cart.RemoveAll(Function(item) item.MaSP = code)
        RefreshCart()
    End Sub

    Private Sub Checkout()
        If _employeeSelect.SelectedValue Is Nothing Then Throw New InvalidOperationException("Cần tạo và chọn nhân viên trước khi thanh toán.")
        Dim employeeId = Convert.ToInt32(_employeeSelect.SelectedValue)
        Dim customerId As Integer? = Nothing
        If _customerSelect.SelectedValue IsNot Nothing AndAlso Convert.ToInt32(_customerSelect.SelectedValue) > 0 Then
            customerId = Convert.ToInt32(_customerSelect.SelectedValue)
        End If

        Dim result = _service.Checkout(employeeId, customerId, _cart)
        MessageBox.Show($"Thanh toán thành công. Hóa đơn: {result.MaHD}{Environment.NewLine}Tổng tiền: {result.TongTien:N0}{Environment.NewLine}Điểm tích lũy: {result.DiemCong}", Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        _cart.Clear()
        RefreshCart()
        RefreshCustomers()
        RefreshSalesLookups()
        RefreshProductsAll()
    End Sub

    Private Sub ClearCart()
        _cart.Clear()
        RefreshCart()
    End Sub

    Private Sub ShowDailyRevenue()
        Dim total = _service.GetDailyRevenue(_dailyDate.Value)
        _reportResult.Text = $"Doanh thu ngày {_dailyDate.Value:yyyy-MM-dd}: {total:N0}"
    End Sub

    Private Sub ShowMonthlyRevenue()
        Dim total = _service.GetMonthlyRevenue(_monthlyDate.Value)
        _reportResult.Text = $"Doanh thu tháng {_monthlyDate.Value:yyyy-MM}: {total:N0}"
    End Sub

    Private Sub ProductSelectionChanged(sender As Object, e As EventArgs)
        If _productGrid.CurrentRow Is Nothing Then Return
        Dim product = TryCast(_productGrid.CurrentRow.DataBoundItem, ProductDto)
        If product Is Nothing Then Return
        _productId.Text = product.MaSP
        _productName.Text = product.TenSP
        _productPrice.Value = Math.Min(_productPrice.Maximum, product.GiaBan)
        _productStock.Value = Math.Min(_productStock.Maximum, product.SoLuongTon)
    End Sub

    Private Sub CustomerSelectionChanged(sender As Object, e As EventArgs)
        If _customerGrid.CurrentRow Is Nothing Then Return
        Dim customer = TryCast(_customerGrid.CurrentRow.DataBoundItem, CustomerDto)
        If customer Is Nothing Then Return
        _selectedCustomerId = customer.MaKH
        _customerName.Text = customer.TenKH
        _customerPhone.Text = customer.DienThoai
    End Sub

    Private Sub EmployeeSelectionChanged(sender As Object, e As EventArgs)
        If _employeeGrid.CurrentRow Is Nothing Then Return
        Dim employee = TryCast(_employeeGrid.CurrentRow.DataBoundItem, EmployeeDto)
        If employee Is Nothing Then Return
        _selectedEmployeeId = employee.MaNV
        _employeeName.Text = employee.TenNV
        _employeePhone.Text = employee.SoDienThoai
        _employeeRole.Text = employee.ChucVu
    End Sub

    Private Sub ClearProductFields()
        _productId.Clear()
        _productName.Clear()
        _productPrice.Value = 0D
        _productStock.Value = 0D
        _productGrid.ClearSelection()
    End Sub

    Private Sub ClearCustomerFields()
        _selectedCustomerId = 0
        _customerName.Clear()
        _customerPhone.Clear()
        _customerGrid.ClearSelection()
    End Sub

    Private Sub ClearEmployeeFields()
        _selectedEmployeeId = 0
        _employeeName.Clear()
        _employeePhone.Clear()
        _employeeRole.Clear()
        _employeeGrid.ClearSelection()
    End Sub

    Private Shared Sub ConfigureNumeric(control As NumericUpDown, maximum As Decimal, decimalPlaces As Integer)
        control.Minimum = 0D
        control.Maximum = maximum
        control.DecimalPlaces = decimalPlaces
        control.Width = If(decimalPlaces = 0, 100, 145)
        control.ThousandsSeparator = True
    End Sub

    Private Shared Sub AddLabeledControl(panel As FlowLayoutPanel, labelText As String, control As Control)
        panel.Controls.Add(New Label With {.Text = labelText, .AutoSize = True, .Margin = New Padding(3, 8, 3, 3)})
        panel.Controls.Add(control)
    End Sub

    Private Shared Function NewButton(caption As String, handler As Action) As Button
        Dim button As New Button With {.Text = caption, .AutoSize = True, .Height = 28, .Margin = New Padding(4)}
        AddHandler button.Click, Sub() handler()
        Return button
    End Function

    Private Shared Function CreateGrid() As DataGridView
        Return New DataGridView With {
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False,
            .RowHeadersVisible = False
        }
    End Function

    Private Shared Sub RunUiAction(action As Action)
        Try
            action()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
