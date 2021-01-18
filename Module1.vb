Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Text

Imports System.Data.OleDb
Imports System.Data.SqlClient

Imports MySql.Data.MySqlClient




Module Module1
	Declare Function GetCurrentTime Lib "kernel32"  Alias "GetTickCount"() As Integer

    Public conn As MySqlConnection
    Public sqlDT As New DataTable
    Public gUser, gPass, gIP, gDBName As String


    Public SQLDT2 As DataTable
	Public Gdb As New ADODB.Connection
	Public gDir As String 'directoy με αρχεία d:\lageuro
	Public gConnect As String 'dbase IV;  Access;
    Dim data As DataTable
    Dim da As MySqlDataAdapter
    Public GdbS As New ADODB.Connection





    Public Function UnicodeBytesToString(ByVal bytes() As Byte) As String
        '// The encoding.
        Dim ascii As ASCIIEncoding = New ASCIIEncoding()

        Return System.Text.Encoding.Unicode.GetString(bytes)
    End Function

    Public Function utftoascii(ByVal unicodeString As String) As String
        ' Dim unicodeString As String = "This string contains the unicode character Pi (" & ChrW(&H3A0) & ")"

        ' Create two different encodings.
        Dim ascii As Encoding = Encoding.ASCII
        Dim unicode As Encoding = Encoding.Unicode

        ' Convert the string into a byte array.
        Dim unicodeBytes As Byte() = unicode.GetBytes(unicodeString)

        ' Perform the conversion from one encoding to the other.
        Dim asciiBytes As Byte() = Encoding.Convert(unicode, ascii, unicodeBytes)

        ' Convert the new byte array into a char array and then into a string.
        Dim asciiChars(ascii.GetCharCount(asciiBytes, 0, asciiBytes.Length) - 1) As Char
        ascii.GetChars(asciiBytes, 0, asciiBytes.Length, asciiChars, 0)
        Dim asciiString As New String(asciiChars)

        ' Display the strings created before and after the conversion.
        Console.WriteLine("Original string: {0}", unicodeString)
        Console.WriteLine("Ascii converted string: {0}", asciiString)
        Return asciiString
    End Function



    Public Sub ExecuteSQLQuery(ByVal SQLQuery As String, ByRef SQLDT2 As DataTable)



        SQLDT2 = New DataTable

        da = New MySqlDataAdapter(SQLQuery, conn)
        ' cb = New MySqlCommandBuilder(da)

        da.Fill(SQLDT2)

        Exit Sub





        'αν χρησιμοποιώ  byref  tote prepei να δηλωθεί   
        'Dim DTI As New DataTable
        SQLDT2 = New DataTable
        'SqlDa=New  
        Dim connStr As String
        'ok working 22-7-2018
        '   connStr = String.Format("pROVIDER=oledb;server={0};user id={1}; password={2}; database=web88_youdb; pooling=false", _
        ' "185.4.134.44", "web88_youdbuser", "youhou!@#$")

        connStr = String.Format("pROVIDER=oledb;server={0};user id={1}; password={2}; database=" + gDBName + "; pooling=false", _
           gIP, gUser, gPass)




        '       connStr = String.Format("server={0};user id={1}; password={2}; database=netbox_data; pooling=false", _
        '"88.99.149.28", "netbox_user", "Wj$7W#ozhLSY")


        '    connStr = String.Format("server={0};user id={1}; password={2}; database=soon_data; pooling=false", _
        '"88.99.149.28", "soon_user", "1o)Nmm!X=P@=")






        Try
            Dim sqlCon As New OleDbConnection(connStr)

            Dim sqlDA As New OleDbDataAdapter(SQLQuery, sqlCon)

            Dim sqlCB As New OleDbCommandBuilder(sqlDA)
            SQLDT2.Reset() ' refresh 
            sqlDA.Fill(SQLDT2)
            ' sqlDA.Fill(sqlDaTaSet, "PEL")

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString)
            If Err.Number = 5 Then
                MsgBox("Invalid Database, Configure TCP/IP", MsgBoxStyle.Exclamation, "Sales and Inventory")
            Else
                MsgBox("Error : " & ex.Message)
            End If
            MsgBox("Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first", MsgBoxStyle.Critical, "Sales And Inventory")
            MsgBox(SQLQuery)
        End Try
        'Return sqlDT
    End Sub


    Public Function ExecuteSQLQuery(ByVal SQLQuery As String) As DataTable
        Try

            Dim connStr As String
            '     connStr = String.Format("PROVIDER=oledb;server={0};user id={1}; password={2}; database=web88_youdb; pooling=false", _
            ' "185.4.134.44", "web88_youdbuser", "youhou!@#$")
            connStr = String.Format("pROVIDER=oledb;server={0};user id={1}; password={2}; database=" + gDBName + "; pooling=false", _
          gIP, gUser, gPass)



            Dim sqlCon As New OleDbConnection(connStr)

            Dim sqlDA As New OleDbDataAdapter(SQLQuery, sqlCon)

            Dim sqlCB As New OleDbCommandBuilder(sqlDA)
            sqlDT.Reset() ' refresh 
            sqlDA.Fill(sqlDT)
            'rowsAffected = command.ExecuteNonQuery();
            ' sqlDA.Fill(sqlDaTaSet, "PEL")

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString)
            If Err.Number = 5 Then
                MsgBox("Invalid Database, Configure TCP/IP", MsgBoxStyle.Exclamation, "Sales and Inventory")
                End

            Else
                MsgBox("Error : " & ex.Message)
            End If
            MsgBox("Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first", MsgBoxStyle.Critical, "Sales And Inventory")
            MsgBox(SQLQuery)
        End Try
        Return sqlDT


    End Function


End Module