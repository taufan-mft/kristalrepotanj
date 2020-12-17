Public Class Form2
    Dim BR_Generator As New MessagingToolkit.Barcode.BarcodeEncoder
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim cek As String = "select * from iphone where ID_phone = '" + TextBox1.Text + "'"
        CMD = New OleDb.OleDbCommand(cek, Conn)


        DM = CMD.ExecuteReader()

        If DM.HasRows = True Then

            'MsgBox("Dianis")
            While DM.Read

                'MsgBox(DM.GetString(0))
                ''Label3.Text = DM.GetString(0)
                Dim total As Single
                total = CInt(TextBox2.Text) * CInt(DM.GetString(5))
                DataGridView1.Rows.Add(New String() {DM.GetString(0), DM.GetString(2), TextBox2.Text, total.ToString})

            End While
        End If
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksiDB()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim total As Single = 0
        For Each row As DataGridViewRow In DataGridView1.Rows
            total = total + (CInt(row.Cells(2).Value) * CInt(row.Cells(3).Value))
        Next
        Debug.WriteLine(total)

        Dim kikyy As String
        kikyy = "C:\Users\Taufan\source\repos\kristalrepotanj\kristalrepotanj"
        kikyy = kikyy + TextBox3.Text + ".jpg"
        PictureBox1.Image.Save(kikyy)
        simpanData("trx", TextBox3.Text, "12/4/2020", total.ToString, kikyy)
        For Each row As DataGridViewRow In DataGridView1.Rows
            If (row.Cells(2).Value Is Nothing) Then
                Exit For
            End If

            Dim kodenya As Integer
            While True
                kodenya = GetRandom(5555, 9999)
                If checkDuplicate("trx_detail", "ID_trxdetail", kodenya.ToString) <> True Then
                    Exit While
                End If
            End While
            simpanData("trx_detail", kodenya.ToString, row.Cells(0).Value, row.Cells(2).Value, row.Cells(3).Value, TextBox3.Text)
        Next
        trxnota.CrystalReport11.SetParameterValue("idtrx", TextBox3.Text)
        trxnota.Show()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Try
            'PictureBox1.Image = QR_Generator.Encode(txtKode.Text)
            PictureBox1.Image = BR_Generator.Encode(MessagingToolkit.Barcode.BarcodeFormat.PDF417, TextBox3.Text)

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
End Class