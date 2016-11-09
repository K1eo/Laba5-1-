Imports System.Data.OleDb


Public Class Access
    Public Structure straight
        Dim index, A, B, C As Integer
        Dim Name, color_straight As String
    End Structure

    Dim names As New List(Of straight)


    Private Sub Add1_Click(sender As Object, e As EventArgs) Handles Add1.Click
        FAdd.ShowDialog()
        Dim lain As String
        Dim name, color, count As String
        count = 1
        Dim A1, B1, C1 As Integer
        Dim res As DialogResult
        A1 = FAdd.TB1.Text
        B1 = FAdd.TB2.Text
        C1 = FAdd.TB3.Text
        name = FAdd.TB4.Text
        color = FAdd.TB5.Text
        res = FAdd.DialogResult
        FAdd.Close()

        If res <> DialogResult.OK Then
            Exit Sub
        End If

        Dim c As New OleDbCommand
        c.Connection = conn
        Dim coefficients As New OleDbCommand("select i from Пряма_у_просторі", conn)
        Dim dr_coefficient As OleDbDataReader = coefficients.ExecuteReader
        While dr_coefficient.Read
            count += 1
        End While
        lain = Convert.ToString(A1) + " " + Convert.ToString(B1) + " " + Convert.ToString(C1)
        c.CommandText = "insert into Пряма_у_просторі (coefficients_ABC,Name,color_straight,i) values('" & lain & "','" & name & "','" & color & "', '" & count & "')"
        c.ExecuteNonQuery()

        refGrip()
    End Sub

    Private Sub Open1_Click(sender As Object, e As EventArgs) Handles Open1.Click
        refGrip()
    End Sub
    ' оновлення на формі БАЗИ ДАНИХ
    Private Sub refGrip()
        Dim c As New OleDbCommand
        c.Connection = conn                                  ' зєднуємось з базою даних
        c.CommandText = "select * from Пряма_у_просторі"      ' задаєм з якої таблиці в базі даних будемо зчитувати інформацію


        Dim ds As New DataSet                                 ' обєкт який хранить n-кількість таблиці в цей обєкт ми занисим певну таблицю яку вказали в ConnetcionText
        Dim da As New OleDbDataAdapter(c)                     '
        da.Fill(ds, "Пряма_у_просторі")
        DataGrid1.DataSource = ds
        DataGrid1.DataMember = "Пряма_у_просторі"

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        conn = New OleDbConnection
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Валентин\Documents\BazaStraight.accdb;Persist Security Info=False;" 'задаємо параметри і адрес до бази даних
        conn.Open()                 'зчитуємо з бази даних
    End Sub

    Private Sub Delete1_Click(sender As Object, e As EventArgs) Handles Delete1.Click
        Dim index As Integer
        Dim c As New OleDbCommand
        c.Connection = conn
        index = DataGrid1.CurrentRow.Cells("i").Value
        c.CommandText = "delete from Пряма_у_просторі where i = " & index
        c.ExecuteNonQuery()
        refGrip()
    End Sub

    ' Sort ///////////////////////////////////////////////////////////////////

    ' файл для відкривання Access і занесення бази даних в conn
    Public Sub outFille()
        conn = New OleDbConnection
        Dim c As New OleDbCommand
        Try
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Валентин\Documents\BazaStraight.accdb;Persist Security Info=False;" 'задаємо параметри і адрес до бази даних
            conn.Open()                 'зчитуємо з бази даних
        Catch ex As Exception
            Console.WriteLine("База даних не знайдена")
        End Try
    End Sub

    Public Sub Input()
        Dim ins As Integer
        Dim str As String
        Dim split1() As String
        ins = 1
        names.RemoveRange(0, names.Count)
        Dim n As New List(Of straight)
        Dim r As New straight

        outFille()
        Dim coefficients As New OleDbCommand("select i,coefficients_ABC,Name,color_straight from Пряма_у_просторі", conn)
        Dim dr_coefficient As OleDbDataReader = coefficients.ExecuteReader
        While (dr_coefficient.Read)
            Try
                r.index = ins
                str = dr_coefficient.Item("coefficients_ABC")
                split1 = Split(str, " ")
                r.A = split1(0)
                r.B = split1(1)
                r.C = split1(2)
                r.Name = dr_coefficient.Item("Name")
                r.color_straight = dr_coefficient.Item("color_straight")
                ins += 1
                names.Add(r)
            Catch ex As Exception
            End Try
        End While

    End Sub

    'додавання в Access базу даних
    Public Sub AddTable()

        Dim lain As String
        Dim i, index As Integer
        i = 0
        index = 1
        Dim coefficients As New OleDbCommand("select i from Пряма_у_просторі", conn)
        Dim dr_coefficient As OleDbDataReader = coefficients.ExecuteReader
        RemovTable()
        While i < names.Count
            Dim c As New OleDbCommand
            c.Connection = conn
            lain = Convert.ToString(names(i).A) + " " + Convert.ToString(names(i).B) + " " + Convert.ToString(names(i).C)
            c.CommandText = "insert into Пряма_у_просторі (i,coefficients_ABC,Name,color_straight) values('" & index & "','" & lain & "','" & names(i).Name & "', '" & names(i).color_straight & "')"
            c.ExecuteNonQuery()
            i += 1
            index += 1
        End While

    End Sub

    ' видалення усієї бази даних Access
    Public Sub RemovTable()
        Dim c As New OleDbCommand
        c.Connection = conn
        Dim coefficients As New OleDbCommand("select i from Пряма_у_просторі", conn)
        Dim dr_coefficient As OleDbDataReader = coefficients.ExecuteReader
        While dr_coefficient.Read
            c.CommandText = "delete from Пряма_у_просторі where i = " & dr_coefficient.Item("i")
            c.ExecuteNonQuery()
        End While
    End Sub

    ' сорутвання бульбашкою =)
    Public Sub bubbleSort()

        Dim n As New straight
        Dim i, j, flag As Integer
        flag = 1
        Console.WriteLine(names.Count)
        While flag > 0
            flag = 0

            For j = 1 To names.Count - 1
                If names(j - 1).Name > names(j).Name Then
                    n.index = names(j - 1).index
                    n.A = names(j - 1).A
                    n.B = names(j - 1).B
                    n.C = names(j - 1).C
                    n.Name = names(j - 1).Name
                    n.color_straight = names(j - 1).color_straight
                    names(j - 1) = names(j)
                    names(j) = n
                    flag = 1
                End If

            Next

        End While
    End Sub

    ' Sort //////////////////////////////////////////////////////////////////////////

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Input()
        bubbleSort()
        AddTable()
    End Sub
End Class
