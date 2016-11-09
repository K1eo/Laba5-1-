Public Class FAdd

    Dim color As New List(Of String)
    Dim color1 As New Hashtable()

    Private Sub UnitializeColor()
        color1.Add("white", 1)
        color1.Add("yellow", 2)
        color1.Add("orange", 3)
        color1.Add("pink", 4)
        color1.Add("red", 5)
        color1.Add("brown", 6)
        color1.Add("green", 7)
        color1.Add("blue", 8)
        color1.Add("purple", 9)
        color1.Add("black", 10)

    End Sub
    ' перша буква велика потім всі маленькі 
    Private Sub Add2_Click(sender As Object, e As EventArgs) Handles Add2.Click
        Me.DialogResult = DialogResult.OK
        Dim str, str1, str2 As String
        str = TB4.Text
        str1 = str.Substring(0, 1) ' беремо першу букву так як вони всі з регістром (ToUpper) 
        str2 = str.Substring(1, TB4.Text.Length - 1).ToLower 'беремо всі крім першої і роюимо маленькими за допомогою (ToLower)
        TB4.Text = str1 + str2
        Me.Close()
    End Sub

    Private Sub Remove2_Click(sender As Object, e As EventArgs) Handles Remove2.Click
        Me.DialogResult = DialogResult.Abort
        Me.Close()
    End Sub

    Private Sub FAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        UnitializeColor()
        Me.AcceptButton = Add2
        Me.CancelButton = Remove2
    End Sub

    Private Sub TB1_TextChanged_KeyPress(sender As Object, e As EventArgs) Handles TB1.TextChanged


    End Sub


    Private Sub Coefficients_A_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TB1.KeyPress
        If e.KeyChar = vbBack Then Exit Sub
        If e.KeyChar < "1" Or e.KeyChar > "9" Then e.Handled = True
    End Sub

    Private Sub Coefficients_B_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TB2.KeyPress
        If e.KeyChar = vbBack Then Exit Sub
        If e.KeyChar < "1" Or e.KeyChar > "9" Then e.Handled = True
    End Sub

    Private Sub Coefficients_C_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TB3.KeyPress
        If e.KeyChar = vbBack Then Exit Sub
        If e.KeyChar < "1" Or e.KeyChar > "9" Then e.Handled = True
    End Sub

    Private Sub textName(sender As Object, e As KeyPressEventArgs) Handles TB4.KeyPress
        If e.KeyChar = vbBack Then Exit Sub
        e.KeyChar = e.KeyChar.ToString.ToUpper
        If Not (e.KeyChar Like "[A-ZA-ЯЇІЄ’]") Then
            e.Handled = True
            Beep()
        End If
    End Sub



    Private Sub FAdd_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
        Dim txt As Control
        For Each txt In Controls
            If TypeOf txt Is TextBox Then
                If txt.Text.Trim = String.Empty Then
                    Add2.Enabled = False
                    Exit Sub
                End If
            End If
        Next
        Add2.Enabled = True
    End Sub

    Private Sub Color_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TB5.Validating
        Dim i As Integer
        i = 0
        If Not color1.ContainsKey(TB5.Text) = True Then
            e.Cancel = True
            MsgBox("колір не існує")
        End If
    End Sub


End Class