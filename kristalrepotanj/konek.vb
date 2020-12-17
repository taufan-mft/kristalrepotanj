Imports System.Data.OleDb



Module konek
    Public Conn As OleDbConnection
    Public DA As OleDbDataAdapter
    Public DS As DataSet
    Public CMD As OleDbCommand
    Public DM As OleDbDataReader

    Sub koneksiDB()
        Try
            Conn = New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source = Database4.accdb")
            Conn.Open()
            'MsgBox("Mantap pan")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub simpanData(namatabel As String, ParamArray var() As TextBox)
        Dim sql As String = "insert into " + namatabel + " values("
        For i As Integer = 0 To UBound(var, 1)
            If i <> UBound(var, 1) Then
                sql = sql + "'" + var(i).Text + "',"
            Else
                sql = sql + "'" + var(i).Text + "')"
            End If

        Next
        CMD = New OleDb.OleDbCommand(sql, Conn)
        CMD.ExecuteNonQuery()
        'MsgBox("BERHAASIL")


    End Sub

    Sub simpanData(namatabel As String, ParamArray var() As String)
        Dim sql As String = "insert into " + namatabel + " values("
        For i As Integer = 0 To UBound(var, 1)
            If i <> UBound(var, 1) Then
                sql = sql + "'" + var(i) + "',"
            Else
                sql = sql + "'" + var(i) + "')"
            End If

        Next
        CMD = New OleDb.OleDbCommand(sql, Conn)
        CMD.ExecuteNonQuery()
        'MsgBox("BERHAASIL")


    End Sub
    Function GenerateRandomString(ByRef iLength As Integer) As String
        Dim rdm As New Random()
        Dim allowChrs() As Char = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLOMNOPQRSTUVWXYZ0123456789".ToCharArray()
        Dim sResult As String = ""

        For i As Integer = 0 To iLength - 1
            sResult += allowChrs(rdm.Next(0, allowChrs.Length))
        Next

        Return sResult
    End Function

    Public Function GetRandom(ByVal Min As Integer, ByVal Max As Integer) As Integer
        Dim Generator As System.Random = New System.Random()
        Return Generator.Next(Min, Max)
    End Function
    Function checkDuplicate(namatabel As String, namaid As String, idkonten As String)
        Dim sequel As String
        sequel = "select * from " + namatabel + " where " + namaid + " = '" + idkonten + "'"
        CMD = New OleDb.OleDbCommand(sequel, Conn)

        DM = CMD.ExecuteReader()
        DM.Read()

        If Not DM.HasRows Then
            Return False
        Else
            Return True
        End If

    End Function
End Module
