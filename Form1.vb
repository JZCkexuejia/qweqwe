Imports System.Data.SqlClient

Public Class Form1
    ' 数据库连接对象
    Private connection As SqlConnection

    ' 窗体加载的时候，创建并保存数据库连接，以及载入数据到 DataGrid
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 定义连接字符串
        Dim connectionString As String = "Server=localhost;Database=neo_db;Integrated Security=True"
        ' 初始化数据库连接对象
        connection = New SqlConnection(connectionString)
        Try
            ' 打开连接
            connection.Open()
            Trace.WriteLine("连接成功！")

            ' 刷新 DataGrid 中数据
            RefreshData()
        Catch ex As Exception
            ' 捕获并显示错误
            Trace.WriteLine("连接失败：" & ex.Message)
        End Try
    End Sub

    ' 在这里获取数据库数据，然后赋值给 DataGridView
    ' 要刷新的时候，也直接调用这个函数就行
    Private Sub RefreshData()
        Dim dataTable As New DataTable() ' 这个写成类成员，不用每次都创建，直接更新数据就可以了，节省资源
        Dim adapter As New SqlDataAdapter("SELECT * FROM dbo.Students", connection)
        adapter.Fill(dataTable)
        DataGridView1.DataSource = dataTable
    End Sub

    Private Sub Button_Add_ROW_Click(sender As Object, e As EventArgs) Handles Button_Add_COL.Click
        Dim name As String = TextBox_Name.Text

        ' 验证列名是否只包含字母和数字
        If Not System.Text.RegularExpressions.Regex.IsMatch(name, "^[a-zA-Z0-9]+$") Then
            Trace.WriteLine("无效的列名")
            Return
        End If

        ' 使用参数化查询来检查列是否存在
        Dim checkColumnSql As String = "
        SELECT COUNT(*) 
        FROM INFORMATION_SCHEMA.COLUMNS 
        WHERE TABLE_NAME = 'Students' AND COLUMN_NAME = @columnName"

        Dim checkCommand As New SqlCommand(checkColumnSql, connection)
        ' 参数不是直接添加到sql字符串那里是为了防止sql注入，不过在这里好像没什么好防的
        checkCommand.Parameters.AddWithValue("@columnName", name)

        Try
            ' 检查列是否存在
            Dim columnExists As Integer = checkCommand.ExecuteScalar()

            If columnExists = 0 Then
                ' 构建安全的添加列的SQL语句
                Dim addColumnSql As String = $"ALTER TABLE dbo.Students ADD [{name}] NVARCHAR(100);"
                Dim addCommand As New SqlCommand(addColumnSql, connection)
                addCommand.ExecuteNonQuery()
            Else
                ' 列存在时，更新列的所有值为 NULL
                Dim updateColumnSql As String = $"UPDATE dbo.Students SET [{name}] = NULL;"
                Dim updateCommand As New SqlCommand(updateColumnSql, connection)
                updateCommand.ExecuteNonQuery()
            End If
            RefreshData()
        Catch ex As Exception
            Trace.WriteLine("添加失败：" & ex.Message)
        End Try
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        ' 窗口关闭时确保连接关闭
        If connection IsNot Nothing AndAlso connection.State = ConnectionState.Open Then
            connection.Close()
        End If
    End Sub

    Private Sub Button_Add_Row_Click_1(sender As Object, e As EventArgs) Handles Button_Add_Row.Click
        Dim name As String = TextBox_Name.Text

        ' 使用参数化查询来检查列是否存在
        Dim checkRowSql As String = "SELECT * FROM dbo.Students WHERE name=@name"

        Dim checkCommand As New SqlCommand(checkRowSql, connection)
        checkCommand.Parameters.AddWithValue("@name", name)

        Try
            ' 检查行是否存在
            Dim rowExists As Integer = checkCommand.ExecuteScalar()

            ' 行不存在时，添加新行
            If rowExists = 0 Then
                Dim addRowSql As String = $"INSERT INTO dbo.Students (name) VALUES (@name)"
                Dim addCommand As New SqlCommand(addRowSql, connection)
                addCommand.Parameters.AddWithValue("@name", name)
                addCommand.ExecuteNonQuery()

                ' 行存在时，更新行的所有值为 NULL
            Else
                '获取所有的列名称
                Dim getColumnSql As String = $"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Students'"
                Dim getColumnNamesCommand As New SqlCommand(getColumnSql, connection)

                ' updateColumns 储存要更新的列
                Dim updateColumns As New List(Of String)
                Using reader As SqlDataReader = getColumnNamesCommand.ExecuteReader()
                    While reader.Read()
                        ' 排除 id 和 name, 排除的原因：id 是行的唯一标识，一般是数据库自动生成的；name 是用户要添加的，不能设为 NULL
                        If reader("COLUMN_NAME") <> "id" And reader("COLUMN_NAME") <> "name" Then
                            ' 添加要更新的列到 updataColumns
                            updateColumns.Add($"[{reader("COLUMN_NAME")}] = NULL")
                        End If
                    End While
                End Using
                ' updateColums 类似这样： ["[test1] = NULL"，"[age] = NULL"]

                Dim updateColumnSql As String = $"UPDATE dbo.Students SET "
                updateColumnSql &= String.Join(", ", updateColumns) & " WHERE [name] = @name;"
                ' updateColumnSql 示例：UPDATE dbo.Students SET [age] = NULL, [test1] = NULL WHERE [name] = @name;

                Dim updateCommand As New SqlCommand(updateColumnSql, connection)
                updateCommand.Parameters.AddWithValue("@name", name)
                updateCommand.ExecuteNonQuery()
            End If
            RefreshData()
        Catch ex As Exception
            Trace.WriteLine("添加失败：" & ex.Message)
        End Try
    End Sub

    Private Sub Button_Refresh_Click(sender As Object, e As EventArgs) Handles Button_Refresh.Click
        RefreshData()
    End Sub
End Class
