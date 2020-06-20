Imports System.IO
Imports System.Security.Cryptography
Imports MySql.Data.MySqlClient


Module Connection
    Public conn As New MySqlConnection 'for MySQLDatabase Connection
    Public msda As MySqlDataAdapter 'is use to update the dataset and datasource
    Public com As New MySqlCommand
    Public rst As MySqlDataReader

    Public file_conn As String = Application.StartupPath.ToString & "\" & My.Application.Info.AssemblyName & ".conn"

    Public sqlserver As String
    Public sqlPort As String
    Public sqluser As String
    Public sqlpass As String
    Public sqldatabase As String

    Public connclient As New MySqlConnection 'for MySQLDatabase Connection
    Public msdaclient As MySqlDataAdapter 'is use to update the dataset and datasource
    Public dstclient As New DataSet 'miniature of your table - cache table to client
    Public comclient As New MySqlCommand
    Public rstclient As MySqlDataReader

    Public clientserver As String
    Public clientport As String
    Public clientuser As String
    Public clientpass As String
    Public clientdatabase As String

    Public GlobalProvinceCode As String
    Public GlobalMunicipalityCode As String
    Public GlobalBarangayCode As String

    Public Function OpenMysqlConnection() As Boolean
        Dim strSetup As String = ""
        Dim sr As StreamReader = File.OpenText(file_conn)
        Dim br As String = sr.ReadLine() : sr.Close()
        strSetup = DecryptTripleDES(br) : Dim cnt As Integer = 0
        For Each word In strSetup.Split(New Char() {","c})
            If cnt = 0 Then
                sqlserver = word
            ElseIf cnt = 1 Then
                sqlPort = word
            ElseIf cnt = 2 Then
                sqluser = word
            ElseIf cnt = 3 Then
                sqlpass = word
            ElseIf cnt = 4 Then
                sqldatabase = word
            End If
            cnt = cnt + 1
        Next
        Try

            conn.Close()
            conn = New MySql.Data.MySqlClient.MySqlConnection
            conn.ConnectionString = "server=" & sqlserver & "; Port=" & sqlPort & "; user id=" & sqluser & "; password=" & sqlpass & "; database=" & sqldatabase & "; Connection Timeout=6000000 ; Allow Zero Datetime=True"
            conn.Open()
            com.Connection = conn
            com.CommandTimeout = 6000000
            LoadDatabaseSettings()
        Catch errMYSQL As MySqlException
            OpenMysqlConnection = False
            Return False
        End Try
        Return True
    End Function

    Public Function OpenClientServer() As Boolean
        Try
            connclient = New MySql.Data.MySqlClient.MySqlConnection
            connclient.ConnectionString = "server=" & clientserver & "; Port=" & clientport & "; user id=" & clientuser & "; password=" & clientpass & "; database=" & clientdatabase & ""
            connclient.Open()
            comclient.Connection = connclient
            comclient.CommandTimeout = 0
            OpenClientServer = True

        Catch errMYSQL As MySqlException
            MessageBox.Show("Message:" & errMYSQL.Message & vbCrLf, _
                             "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            OpenClientServer = False
            Return False
        Catch errMS As Exception
            MessageBox.Show("Message:" & errMS.Message & vbCrLf, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            OpenClientServer = False
            Return False
        End Try
    End Function


    Public Sub LoadDatabaseSettings()
        com.CommandText = "select * from tblgeneralsettings" : rst = com.ExecuteReader
        While rst.Read
            GlobalProvinceCode = rst("province").ToString
            GlobalMunicipalityCode = rst("municipal").ToString
            GlobalBarangayCode = rst("unitcode").ToString
        End While
        rst.Close()
    End Sub

    Const sKey As String = "kira"

    Public Function EncryptTripleDES(ByVal sIn As String) As String
        Dim DES As New TripleDESCryptoServiceProvider()
        Dim hashMD5 As New MD5CryptoServiceProvider()

        ' Compute the MD5 hash.
        DES.Key = hashMD5.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(sKey))
        ' Set the cipher mode.
        DES.Mode = CipherMode.ECB
        ' Create the encryptor.
        Dim DESEncrypt As ICryptoTransform = DES.CreateEncryptor()
        ' Get a byte array of the string.
        Dim Buffer As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(sIn)
        ' Transform and return the string.
        Return Convert.ToBase64String(DESEncrypt.TransformFinalBlock(Buffer, 0, Buffer.Length))
    End Function

    Public Function DecryptTripleDES(ByVal sOut As String) As String
        Dim DES As New TripleDESCryptoServiceProvider()
        Dim hashMD5 As New MD5CryptoServiceProvider()

        ' Compute the MD5 hash.
        DES.Key = hashMD5.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(sKey))
        ' Set the cipher mode.
        DES.Mode = CipherMode.ECB
        ' Create the decryptor.
        Dim DESDecrypt As ICryptoTransform = DES.CreateDecryptor()
        Dim Buffer As Byte() = Convert.FromBase64String(sOut)
        ' Transform and return the string.
        Return System.Text.ASCIIEncoding.ASCII.GetString(DESDecrypt.TransformFinalBlock(Buffer, 0, Buffer.Length))
    End Function
    Public Function rchar(ByVal str As String)
        str = str.Replace("'", "''")
        str = str.Replace("\", "\\")
        str = str.Replace("Ã‘", "Ñ")
        Return str
    End Function
    Public Function countqry(ByVal tbl As String, ByVal cond As String)
        Dim cnt As Integer = 0
        com.CommandText = "select count(*) as cnt from " & tbl & " where " & cond
        rst = com.ExecuteReader
        While rst.Read
            cnt = Val(rst("cnt").ToString)
        End While
        rst.Close()
        Return cnt
    End Function
    Public Function ConvertDate(ByVal d As Date)
        Return d.ToString("yyyy-MM-dd")
    End Function

    Public Function getResidentidSequence()
        Dim strng As Integer = 0 : Dim newNumber As String = "" : Dim NumberLen As Integer = 0
        com.CommandText = "select membersequence from tblgeneralsettings" : rst = com.ExecuteReader()
        While rst.Read
            NumberLen = rst("membersequence").ToString.Length
            strng = Val(rst("membersequence").ToString) + 1
        End While
        rst.Close()
        If NumberLen > strng.ToString.Length Then
            Dim a As Integer = NumberLen - strng.ToString.Length
            If a = 10 Then
                newNumber = "0000000000" & strng
            ElseIf a = 9 Then
                newNumber = "000000000" & strng
            ElseIf a = 8 Then
                newNumber = "00000000" & strng
            ElseIf a = 7 Then
                newNumber = "0000000" & strng
            ElseIf a = 6 Then
                newNumber = "000000" & strng
            ElseIf a = 5 Then
                newNumber = "00000" & strng
            ElseIf a = 4 Then
                newNumber = "0000" & strng
            ElseIf a = 3 Then
                newNumber = "000" & strng
            ElseIf a = 2 Then
                newNumber = "00" & strng
            ElseIf a = 1 Then
                newNumber = "0" & strng
            Else
                newNumber = strng
            End If
        Else
            newNumber = strng
        End If
        com.CommandText = "update tblgeneralsettings set membersequence='" & newNumber & "'" : com.ExecuteNonQuery()
        Return GlobalBarangayCode & "-" & newNumber
    End Function

    Public imgBytes As Byte() = Nothing
    Public stream As MemoryStream = Nothing
    Public img As Image = Nothing
    Public sqlcmd As New MySqlCommand
    Public sql As String
    Public arrImage() As Byte = Nothing

    Public Function UpdateImage(ByVal qry As String, ByVal fld As String, ByVal tbl As String, ByVal picbox As System.Windows.Forms.PictureBox)
        arrImage = Nothing
        Try
            If Not picbox.Image Is Nothing Then
                Dim mstream As New System.IO.MemoryStream()
                picbox.Image.Save(mstream, System.Drawing.Imaging.ImageFormat.Gif)
                arrImage = mstream.GetBuffer()
                mstream.Close()
            End If

            sql = "Update " & tbl & " set " & fld & " = @file where " & qry

            With sqlcmd
                .CommandText = sql
                .Connection = conn
                .Parameters.AddWithValue("@file", arrImage)
                .ExecuteNonQuery()
            End With
            sqlcmd.Parameters.Clear()

            Dim qryTask As String = sqlcmd.CommandText.ToString
            Dim qryStr As String = "INSERT into filedir.`tblactionquerylogs` set dateperformed=current_timestamp, editedtype='Image', querytask='" & rchar(qryTask) & "', remarks='UPDATE IMAGE', performedby='Migrated', img=@file "
            With com
                .CommandText = qryStr
                .Connection = conn
                .Parameters.AddWithValue("@file", arrImage)
                .ExecuteNonQuery()
            End With
            com.Parameters.Clear()

        Catch errMYSQL As MySqlException
            MessageBox.Show("Message:" & errMYSQL.Message & vbCrLf, _
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch errMS As Exception
            MessageBox.Show("Message:" & errMS.Message & vbCrLf, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return 0
    End Function

    Public Function ResizedImage(ByVal img As Image) As Image
        ResizedImage = Nothing
        If img Is Nothing Then Exit Function
        Dim Original As New Bitmap(img)
        img = Original

        Dim m As Int32 = 400
        Dim n As Int32 = m * Original.Height / Original.Width

        Dim Thumb As New Bitmap(m, n, Original.PixelFormat)
        Thumb.SetResolution(Original.HorizontalResolution, Original.VerticalResolution)

        Dim g As Graphics = Graphics.FromImage(Thumb)
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High

        g.DrawImage(Original, New Rectangle(0, 0, m, n))
        ResizedImage = Thumb
        Return ResizedImage
    End Function

    Public Function LogQuery(ByVal editedtype As String, ByVal querytask As String, ByVal remarks As String)
        com.CommandText = "insert into filedir.tblactionquerylogs set dateperformed=current_timestamp, editedtype='" & rchar(editedtype) & "', querytask='" & rchar(querytask) & "', remarks='" & rchar(remarks) & "', performedby='Migrated'" : com.ExecuteNonQuery()
        Return 0
    End Function

End Module
