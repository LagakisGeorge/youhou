Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.Office.Interop
Imports MySql.Data.MySqlClient

Friend Class Form1
    Inherits System.Windows.Forms.Form

    Dim MT As DataTable





    Dim data As DataTable
    Dim da As MySqlDataAdapter
    Dim cb As MySqlCommandBuilder





    'I just solved my problem just like this:
    '
    '1) Connect to the db with mysql workbench (select the "use old auth" and "send pass in cleartext"
    '
    '2) Using the workbench, I ran the following several times (because it kept saying '0 rows affected'): (also change userID and password to your proper stuff)
    '
    'SET SESSION old_passwords=0;
    'SET PASSWORD FOR userID=PASSWORD('password');
    '
    'SET SESSION old_passwords=false;
    'SET PASSWORD FOR userID=PASSWORD('password');
    '3) Now go back to your app and it should run...
    '
    'That is how I did it at least. I am using IX mysql and they do use old auth on their server so you must do something on your end..
    '
    '

    Dim fr As New ADODB.Recordset
    Dim fr2 As New ADODB.Recordset
    Dim mF100 As Integer



    Dim fSQL As Object

    Dim elem(300) As Object
    Dim nEkkr As Object
    Dim link As Object
    Dim fLinks(300) As Object
    Dim gcon2 As String
    Dim SURL As Object
    Dim Ekkremeis(90) As Object
    Dim nCounter As Object ' counter paraggelion

    'UPGRADE_WARNING: Event Check1.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub Check1_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check1.CheckStateChanged
        '<EhHeader>
        On Error GoTo oles_Click_Err
        '</EhHeader>


100:    'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fSQL = "SELECT   *  FROM oc_order ORDER BY order_id  DESC limit  " & Str(mF100)
102:    'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        Adodc1.RecordSource = fSQL
        '104:    Adodc1.Refresh()



        '<EhFooter>
        Exit Sub

oles_Click_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.oles_Click " & "at line " & Erl())
        Resume Next
        '</EhFooter>

    End Sub

    Private Sub cmdEKKA»¡—…”«–—œ”Ÿ—…Õœ’_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEKKA»¡—…”«–—œ”Ÿ—…Õœ’.Click
        Gdb.Execute("DELETE FROM PEGGTIM")
        Gdb.Execute("DELETE FROM PPEL")


    End Sub

    'UPGRADE_WARNING: Event Combo1.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub Combo1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.SelectedIndexChanged
        '<EhHeader>
        On Error GoTo Combo1_Click_Err
        '</EhHeader>

100:    'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fSQL = "SELECT  *  FROM oc_order where order_status_id=" & VB.Left(Combo1.Text, 2) & " ORDER BY date_added DESC limit  " & Str(mF100)
102:    'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Adodc1.RecordSource = fSQL
104:    ' Adodc1.Refresh()

        'Exit Sub


        data = New DataTable

        da = New MySqlDataAdapter(fSQL, conn)
        ' cb = New MySqlCommandBuilder(da)

        da.Fill(data)

        DataGrid1.DataSource = data

        ' DataGrid1.ReBind()
        DataGrid1.Refresh()

        DataGrid1.DataSource = data




        Dim r As New ADODB.Recordset
106:    r.Open("select  count(*) FROM oc_order where order_status_id=" & VB.Left(Combo1.Text, 2), GdbS, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
108:    Me.Text = "–¡—¡√√≈À…≈” " & Str(r.Fields(0).Value)
110:    r.Close()


        '<EhFooter>
        Exit Sub

Combo1_Click_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.Combo1_Click " & "at line " & Erl())
        Resume Next
        '</EhFooter>
    End Sub




    Private Sub connect_MySQL()
        If Not conn Is Nothing Then conn.Close()

        Dim connStr As String
        '       connStr = String.Format("server={0};user id={1}; password={2}; database=web88_youdb; pooling=false", _
        '  " 88.99.239.138", "web88_youdbuser", "youhou!@#$")

        gDBName = "web88_youdb"


        gIP = "185.4.134.29" '88.99.239.138"
        gUser = "web88_youdbuser"
        gPass = "youhou!@#$"



        ' gcon2 = String.Format("server={0};user id={1}; password={2}; database=onemore_mon498; pooling=false", _
        ' "185.4.132.41", "onemore_musr4", "kkaPWoBaPKpwdq5")


        connStr = String.Format("server={0};user id={1}; password={2}; database=" + gDBName + "; pooling=false", _
          gIP, gUser, gPass)

        '    connStr = String.Format("server={0};user id={1}; password={2}; database=netbox_data; pooling=false", _
        ' "88.99.149.28", "netbox_user", "Wj$7W#ozhLSY")


        '  connStr = String.Format("server={0};user id={1}; password={2}; database=soon_data; pooling=false", _
        '  "88.99.149.28", "soon_user", "1o)Nmm!X=P@=")

        ' 2815102911  kodikos aithmatos BER673-29500  AITHMA  
        ' INFO@SWA.GR





        Try
            conn = New MySqlConnection(connStr)
            conn.Open()

            'GetDatabases()
        Catch ex As MySqlException
            MessageBox.Show("Error connecting to the server: " + ex.Message)
        End Try
    End Sub






    Private Sub ALL_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ALL.Click
        '<EhHeader>
        On Error GoTo ALL_Click_Err
        '</EhHeader>
        '100:    Do While Not Adodc1.Recordset.EOF
        '102:        EXEC_Click(EXEC, New System.EventArgs())
        '104:        Adodc1.Recordset.MoveNext()
        '        Loop

106:    MsgBox("œ ")

        '<EhFooter>
        Exit Sub

ALL_Click_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.ALL_Click " & "at line " & Erl())
        Resume Next
        '</EhFooter>
    End Sub

    Private Sub EYRESH_EKKREMON()
        Dim linkparagg As Object
        Dim List1 As Object
        Dim WebBrowser1 As Object
        '<EhHeader>
        On Error GoTo EYRESH_EKKREMON_Err
        '</EhHeader>
        Dim s As Object

        Dim k, nn As Object
100:    For k = 1 To 300
102:        'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object elem(k). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            elem(k) = ""
        Next


104:    'UPGRADE_WARNING: Couldn't resolve default property of object WebBrowser1.document. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        s = WebBrowser1.document.ALL.Item(1).innerHTML
106:    FETES_DELIM(s, elem)

108:    For k = 1 To 300
110:        'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object elem(k). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object List1.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            List1.AddItem(VB6.Format(k, "000") & " " + elem(k))
        Next

112:    'UPGRADE_WARNING: Couldn't resolve default property of object linkparagg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        linkparagg = 35
114:    'UPGRADE_WARNING: Couldn't resolve default property of object nEkkr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nEkkr = 0
116:    For k = 138 To 300 Step 5
118:        'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object elem(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If InStr(elem(k), "ÂÍÍÒÂÏﬁÚ") > 0 Then


120:            'UPGRADE_WARNING: Couldn't resolve default property of object nEkkr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nEkkr = nEkkr + 1
122:            'UPGRADE_WARNING: Couldn't resolve default property of object nEkkr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object linkparagg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Ekkremeis(nEkkr). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Ekkremeis(nEkkr) = linkparagg ' ·Ì ÂÈÌ·È ÂÍÍÒÂÏÁÚ ﬁ ¸˜È

            Else
                '
            End If
124:        'UPGRADE_WARNING: Couldn't resolve default property of object linkparagg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            linkparagg = linkparagg + 1
        Next

126:    'UPGRADE_WARNING: Couldn't resolve default property of object nEkkr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        MsgBox("ÂÍÍÒÂÏÂﬂÚ ·Ò·„„ÂÎﬂÂÚ " & Str(nEkkr))

        '<EhFooter>
        Exit Sub

EYRESH_EKKREMON_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.EYRESH_EKKREMON " & "at line " & Erl())
        Resume Next
        '</EhFooter>
    End Sub
    Private Sub EYRESH2_EKKREMON()
        Dim linkparagg As Object
        Dim List1 As Object
        Dim WebBrowser1 As Object
        Dim List2 As Object
        '<EhHeader>
        On Error GoTo EYRESH2_EKKREMON_Err
        '</EhHeader>
        Dim s As Object
        'Dim links(300) As String

100:    'UPGRADE_WARNING: Couldn't resolve default property of object List2.Clear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        List2.Clear()

        Dim k, nn As Object
102:    For k = 1 To 300
104:        'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object elem(k). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            elem(k) = ""
106:        'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object fLinks(k). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            fLinks(k) = ""
        Next


108:    'UPGRADE_WARNING: Couldn't resolve default property of object WebBrowser1.document. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        s = WebBrowser1.document.ALL.Item(1).innerHTML
110:    FETES_DELIM(s, elem)

112:    For k = 1 To 300
114:        'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object elem(k). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object List1.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            List1.AddItem(VB6.Format(k, "000") & " " + elem(k))
        Next

        Dim allLinks(300) As Object
        Dim allCounter As Object
116:    'UPGRADE_WARNING: Couldn't resolve default property of object allCounter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        allCounter = 1



        'ÂÒ·Ù‹˘ ÂÌ·-ÂÌ· Ù· links Ï›˜ÒÈ Ì· ‚Ò˛ ÙÔ ÎÈÌÍ Ôı ÂÒÈ›˜ÂÈ ·Ò·„„ÂÎﬂ·
        'ÙÔ Í·Ù·Î·‚·ﬂÌ˘ ·¸ ÙÔ
        ' If InStr(s, "action=edit") > 0 And InStr(s, "oID=") > 0 Then
        'Í·È ‚Î›˘ ¸ÙÈ ÙÔ Ò˛ÙÔ ÎÈÌÍ ÏÂ ÙÁÌ ·Ò·„„ÂÎﬂ· ÂﬂÌ·È
        ' ¸Ù·Ì ÙÔ Í=9
        ' „È· Ì· È‹Ì˘ ÙÔ 9 ˜ÒÁÛÈÏÔÔÈ˛ ÙÁÌ ÏÂÙ·‚ÎÁÙﬁ  first_paragg_link

        Dim first_paragg_link As Object
118:    'UPGRADE_WARNING: Couldn't resolve default property of object first_paragg_link. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        first_paragg_link = 0

120:    For k = 1 To 60 '1 to 60
            On Error Resume Next
            'brisko ta links ton paragelion
122:        'UPGRADE_WARNING: Couldn't resolve default property of object WebBrowser1.document. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            s = WebBrowser1.document.links(k)
124:        'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If InStr(s, "action=edit") > 0 And InStr(s, "oID=") > 0 Then
126:            'UPGRADE_WARNING: Couldn't resolve default property of object allCounter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object allLinks(allCounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                allLinks(allCounter) = s ' ÈÌ·Í·Ú ¸Î˘Ì Ù˘Ì ·Ò·„„ÂÎÈ˘Ì ÙÁÚ ÛÂÎﬂ‰·Ú
128:            'UPGRADE_WARNING: Couldn't resolve default property of object allCounter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                allCounter = allCounter + 1
130:            'UPGRADE_WARNING: Couldn't resolve default property of object first_paragg_link. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If first_paragg_link = 0 Then
132:                'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object first_paragg_link. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    first_paragg_link = k 'ÙÛÈÏ‹˘ ÙÔ Ò˛ÙÔ link Ôı ÂÒÈ›˜ÂÈ ·Ò·„„ÂÎﬂÂÚ
                End If
            End If
        Next



        'ÙÔ Ò˛ÙÔ link ÏÂ ÙÁÌ ·Ò·„„ÂÎﬂ· ÂﬂÌ·È ÙÔ 9 ‰ÁÎ·‰ﬁ ÙÔ
        ' WebBrowser1.Document.links(9)   ÓÂÍÈÌ˛ ÏÂ 7 Í·È ÒÔÛËÂÙ˘ +2

134:    'UPGRADE_WARNING: Couldn't resolve default property of object first_paragg_link. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object linkparagg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        linkparagg = first_paragg_link
136:    'UPGRADE_WARNING: Couldn't resolve default property of object allCounter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        allCounter = 0
138:    'UPGRADE_WARNING: Couldn't resolve default property of object nEkkr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nEkkr = 0

        'ÂÒ·Ù‹˘ ÙÈÚ ÛÂÈÒ›Ú ÙÁÚ ÛÂÎﬂ‰·Ú(elem(k))  ·Ì· 6 „È· Ì· ‰˘ ·Ì ı‹Ò˜ÂÈ Á Î›ÓÁ "ÂÍÍÒÂÏﬁÚ"
        '‰ÁÎ·‰ﬁ Á Ò˛ÙÁ ·Ò·„„ÂÎﬂ· ÂﬂÌ·È ÙÔ link(9) Í·È ÙÔ ÛÁÏ‹‰È ·Ì ÂﬂÌ·È
        'ÂÍÍÒÂÏﬁÚ ﬁ Ô˜È ÂÈÌ·È ÛÙÁÌ ÛÂÈÒ‹ 88
        'Print Ekkremeis(1)
        '9
        'Print fLinks(1)
        'http://www.toys-shop.gr/admin/orders.php?page=1&oID=1140&action=edit







        'ÂÒ·Ù‹˘ Ï¸ÌÔ Ù· Û˜¸ÎÈ· Ù˘Ì ·Ò·„„ÂÎÈ˛Ì „È· Ì· ÂÌÙÔﬂÛ˘ ÙÈÚ ÂÍÍÒÂÏÂﬂÚ
140:    For k = 88 To 300 Step 6 ' 122 to 300
142:        'UPGRADE_WARNING: Couldn't resolve default property of object allCounter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            allCounter = allCounter + 1


            On Error Resume Next
144:        'UPGRADE_WARNING: Couldn't resolve default property of object WebBrowser1.document. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object linkparagg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object List2.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            List2.AddItem(CDbl(Str(linkparagg) & " ") + WebBrowser1.document.links(linkparagg))
146:        'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object elem(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If InStr(elem(k), "ÂÍÍÒÂÏﬁÚ") > 0 Then


148:            'UPGRADE_WARNING: Couldn't resolve default property of object nEkkr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nEkkr = nEkkr + 1
150:            'UPGRADE_WARNING: Couldn't resolve default property of object nEkkr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object linkparagg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Ekkremeis(nEkkr). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Ekkremeis(nEkkr) = linkparagg ' ·Ì ÂÈÌ·È ÂÍÍÒÂÏÁÚ ﬁ ¸˜È


                'allLinks ÈÌ·Í·Ú ¸Î˘Ì Ù˘Ì ·Ò·„„ÂÎÈ˘Ì ÙÁÚ ÛÂÎﬂ‰·Ú
                'fLinks(nEkkr) ﬂÌ·Í·Ú Ù˘Ì ≈  —≈ÃŸÕ –¡—¡√√≈À…ŸÕ

152:            'UPGRADE_WARNING: Couldn't resolve default property of object nEkkr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object allCounter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object allLinks(allCounter). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object fLinks(nEkkr). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                fLinks(nEkkr) = allLinks(allCounter) ' WebBrowser1.Document.links(linkparagg)
            Else
                '
            End If
            'If linkparagg = 28 Then linkparagg = linkparagg + 1 Else
154:        'UPGRADE_WARNING: Couldn't resolve default property of object linkparagg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            linkparagg = linkparagg + 2
        Next

156:    'UPGRADE_WARNING: Couldn't resolve default property of object nEkkr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        MsgBox("ÂÍÍÒÂÏÂﬂÚ ·Ò·„„ÂÎﬂÂÚ " & Str(nEkkr))

        '<EhFooter>
        Exit Sub

EYRESH2_EKKREMON_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.EYRESH2_EKKREMON " & "at line " & Erl())
        Resume Next
        '</EhFooter>
    End Sub

    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        '        '<EhHeader>
        '        'On Error GoTo Command1_Click_Err
        '        '</EhHeader>
        '        Dim Excel As Microsoft.Office.Interop.Excel.Application
        '        'UPGRADE_ISSUE: Excel.workbook object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        Dim workbook As Excel.workbook
        '        'UPGRADE_ISSUE: Excel.Worksheet object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        Dim myXL As Excel.Worksheet
        '100:    Excel = CreateObject("excel.Application")
        '102:    workbook = ExcelGlobal_definst.Workbooks.Add
        '        On Error Resume Next
        '        Dim F_XROMATA As Short
        '104:    'UPGRADE_WARNING: Couldn't resolve default property of object workbook.Activate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        workbook.Activate()

        '106:    'UPGRADE_WARNING: Couldn't resolve default property of object workbook.ActiveSheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL = workbook.ActiveSheet




        '        Dim k, rr As Object
        '108:    Adodc1.Recordset.MoveFirst()

        '110:    'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        rr = 1

        '        Dim AA As Short
        '        ' AA = InputBox("1.ACS  2.√≈Õ… « 3.–œ—‘¡-–œ—‘¡ 4.SPEEDEX", , 1)
        '        AA = 3
        '114:    'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
        '        If AA = 2 Then GENIKI()
        '116:    'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
        '        If AA = 1 Then ACS()
        '118:    'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
        '        If AA = 3 Then PORTA()
        '119:    'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
        '        If AA = 4 Then speedex()

        '        '¡CS
        '        'œÕœÃ¡ –¡—¡À«–‘« ≈–ŸÕ’Ã…¡ ≈‘¡…—…¡”   –≈—…œ◊« œƒœ”    ¡—…»Ãœ” œ—œ÷œ”  ‘   Àœ…–¡   ‘«À≈÷ŸÕœ     …Õ«‘œ  ’–œ ¡‘¡”‘«Ã¡    –¡—¡‘«—«”≈…”    ◊—≈Ÿ”«  ‘≈Ã¡◊…¡ ¬¡—œ”   ¡Õ‘… ¡‘¡¬œÀ«    ‘—œ–œ” –À«—ŸÃ«” ¡”÷¡À≈…¡     ≈Õ‘—œ  œ”‘œ’”  ”◊≈‘… œ1    ”◊≈‘… œ2    Ÿ—¡ –¡—¡ƒœ”«”   –—œ…œÕ‘¡
        '        'delivery_name   delivery_company    delivery_city   delivery_street_address ¡ÒÈËÏ¸Ú default=null    ºÒÔˆÔÚ default=0    delivery_postcode   ÀÔÈ‹ default=null  customers_telephone customers_fax   ’ÔÍ·Ù‹ÛÙÁÏ· default=ÃÔı„Ô˝‰ÁÚ …. & ”…¡ œ≈ –·Ò·ÙÁÒﬁÛÂÈÚ default= ÍÂÌ¸  ◊Ò›˘ÛÁ default=A    ‘ÂÏ‹˜È· default=1   ¬‹ÒÔÚ default = 2   if payment_method=¡ÌÙÈÍ·Ù·‚ÔÎﬁ = value  ‘Ò¸ÔÚ ÎÁÒ˘ÏﬁÚ default =M  ¡Ûˆ‹ÎÂÈ· default =null   ÂÌÙ default=null   ”˜ÂÙÈÍ¸ 1 default=orders_id ”˜ÂÙÈÍ¸ 2 default=null  øÒ· –·Ò‹‰ÔÛÁÚ default=null  –ÒÔﬂÔÌÙ· default=null
        '        Dim ANTIK As Single

        '120:    Do While Not Adodc1.Recordset.EOF

        '122:        ANTIK = 0
        '124:        fr2.Open("select * from orders_total where orders_id=" & Str(Adodc1.Recordset.Fields("orders_id").Value) & " and class='ot_fixed_payment_chg'")
        '126:        If Not fr2.EOF Then
        '128:            ANTIK = fr2.Fields("VALUE").Value
        '            End If
        '130:        fr2.Close()

        '132:        For k = 0 To Adodc1.Recordset.Fields.Count - 1
        '134:            'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                myXL.cells(rr, k + 1) = Adodc1.Recordset.Fields(k).Value
        '            Next
        '136:        Adodc1.Recordset.MoveNext()
        '138:        'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            rr = rr + 1

        '        Loop
        '140:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.SaveAs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.SaveAs("c:\parag.xlsx")


        '142:    Excel.Quit()

    End Sub


    Sub ACS()
        '        Dim antik As Integer
        '        '¡CS
        '        'œÕœÃ¡ –¡—¡À«–‘« ≈–ŸÕ’Ã…¡ ≈‘¡…—…¡”   –≈—…œ◊« œƒœ”    ¡—…»Ãœ” œ—œ÷œ”  ‘   Àœ…–¡   ‘«À≈÷ŸÕœ     …Õ«‘œ  ’–œ ¡‘¡”‘«Ã¡    –¡—¡‘«—«”≈…”    ◊—≈Ÿ”«  ‘≈Ã¡◊…¡ ¬¡—œ”   ¡Õ‘… ¡‘¡¬œÀ«    ‘—œ–œ” –À«—ŸÃ«” ¡”÷¡À≈…¡     ≈Õ‘—œ  œ”‘œ’”  ”◊≈‘… œ1    ”◊≈‘… œ2    Ÿ—¡ –¡—¡ƒœ”«”   –—œ…œÕ‘¡
        '        'delivery_name   delivery_company    delivery_city   delivery_street_address ¡ÒÈËÏ¸Ú default=null    ºÒÔˆÔÚ default=0    delivery_postcode   ÀÔÈ‹ default=null  customers_telephone customers_fax   ’ÔÍ·Ù‹ÛÙÁÏ· default=ÃÔı„Ô˝‰ÁÚ …. & ”…¡ œ≈ –·Ò·ÙÁÒﬁÛÂÈÚ default= ÍÂÌ¸  ◊Ò›˘ÛÁ default=A    ‘ÂÏ‹˜È· default=1   ¬‹ÒÔÚ default = 2   if payment_method=¡ÌÙÈÍ·Ù·‚ÔÎﬁ = value  ‘Ò¸ÔÚ ÎÁÒ˘ÏﬁÚ default =M  ¡Ûˆ‹ÎÂÈ· default =null   ÂÌÙ default=null   ”˜ÂÙÈÍ¸ 1 default=orders_id ”˜ÂÙÈÍ¸ 2 default=null  øÒ· –·Ò‹‰ÔÛÁÚ default=null  –ÒÔﬂÔÌÙ· default=null
        '        Dim k, rr As Object


        '        Dim Excel As Microsoft.Office.Interop.Excel.Application
        '        'UPGRADE_ISSUE: Excel.workbook object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        Dim workbook As Excel.Workbook
        '        'UPGRADE_ISSUE: Excel.Worksheet object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        Dim myXL As Excel.Worksheet
        '100:    Excel = CreateObject("excel.Application")
        '102:    workbook = ExcelGlobal_definst.Workbooks.Add
        '        On Error Resume Next
        '        Dim F_XROMATA As Short
        '104:    'UPGRADE_WARNING: Couldn't resolve default property of object workbook.Activate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        workbook.Activate()

        '106:    'UPGRADE_WARNING: Couldn't resolve default property of object workbook.ActiveSheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL = workbook.ActiveSheet





        '144:    'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        rr = 1
        '146:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 1) = "œÕœÃ¡ –¡—¡À«–‘«"
        '148:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 2) = "≈–ŸÕ’Ã…¡ ≈‘¡…—…¡”"
        '150:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 3) = "–≈—…œ◊«"
        '152:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 4) = "œƒœ”"
        '154:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 5) = "¡—…»Ãœ”"
        '156:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 6) = "œ—œ÷œ”"
        '158:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 7) = "‘ "
        '160:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 8) = "Àœ…–¡"
        '162:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 9) = "‘«À≈÷ŸÕœ"
        '164:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 10) = " …Õ«‘œ"
        '166:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 11) = "’–œ ¡‘¡”‘«Ã¡"
        '168:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 12) = "–¡—¡‘«—«”≈…”"
        '170:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 13) = "◊—≈Ÿ”«"
        '172:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 14) = "‘≈Ã¡◊…¡"
        '174:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 15) = "¬¡—œ”"
        '176:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 16) = "¡Õ‘… ¡‘¡¬œÀ«"
        '178:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 17) = "‘—œ–œ” –À«—ŸÃ«”"
        '180:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 18) = "¡”÷¡À≈…¡"
        '182:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 19) = " ≈Õ‘—œ  œ”‘œ’”"
        '184:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 20) = "”◊≈‘… œ1"
        '186:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 21) = "”◊≈‘… œ2"
        '188:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 22) = "Ÿ—¡ –¡—¡ƒœ”«”"
        '190:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Cells(rr, 23) = "–—œ…œÕ‘¡"


        '192:    'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        rr = 2
        '194:    Do While Not Adodc1.Recordset.EOF

        '196:        ANTIK = 0
        '198:        If VB.Left(Adodc1.Recordset.Fields("payment_method").Value, 5) = "¡ÌÙÈÍ" Then

        '200:            fr2.Open("select * from orders_total where orders_id=" & Str(Adodc1.Recordset.Fields("orders_id").Value) & " and class ='ot_total'", GdbS, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        '                '   + " and class IN ('ot_fixed_payment_chg','ot_shipping')"

        '202:            If Not fr2.EOF Then
        '204:                ANTIK = System.Math.Round(fr2.Fields("value").Value, 2)

        '                    'Do While Not fr2.EOF
        '                    '  For k = 0 To fr2.Fields.Count - 1
        '                    '     myXL.cells(rr, k + 1) = fr2(k).Name
        '                    '     myXL.cells(rr + 1, k + 1) = fr2(k)
        '                    '  Next
        '                    '  rr = rr + 2
        '                    '  fr2.MoveNext
        '                    'Loop
        '                End If
        '206:            fr2.Close()
        '            End If



        '208:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 1) = Adodc1.Recordset.Fields("delivery_name").Value
        '210:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 2) = Adodc1.Recordset.Fields("delivery_company").Value
        '212:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 3) = Adodc1.Recordset.Fields("delivery_city").Value
        '214:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 4) = Adodc1.Recordset.Fields("delivery_street_address").Value
        '216:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 5) = "."
        '218:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 6) = 0
        '220:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 7) = Adodc1.Recordset.Fields("delivery_postcode").Value
        '            'myXL.cells(rr, 8) = "Àœ…–¡"
        '222:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 9) = Adodc1.Recordset.Fields("customers_telephone").Value
        '224:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 10) = Adodc1.Recordset.Fields("customers_fax").Value
        '226:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 11) = "ÃÔı„Ô˝‰ÁÚ …. & ”…¡ œ≈"
        '228:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 12) = " "
        '230:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 13) = "A" ' "◊—≈Ÿ”«"
        '232:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 14) = 1 'TEMAXIA
        '234:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 15) = 2 'BAROS



        '236:        If ANTIK > 0 Then
        '238:            'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                myXL.Cells(rr, 16) = VB6.Format(System.Math.Round(ANTIK, 2), "#####.##")
        '            End If

        '240:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 17) = "M" ' "‘—œ–œ” –À«—ŸÃ«”"
        '            'myXL.cells(rr, 18) = "¡”÷¡À≈…¡"
        '            'myXL.cells(rr, 19) = " ≈Õ‘—œ  œ”‘œ’”"
        '242:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.Cells(rr, 20) = Adodc1.Recordset.Fields("orders_id").Value ' ”◊≈‘… œ1"
        '            'myXL.cells(rr, 21) = "”◊≈‘… œ2"
        '            'myXL.cells(rr, 22) = "Ÿ—¡ –¡—¡ƒœ”«”"
        '            'myXL.cells(rr, 23) = "–—œ…œÕ‘¡"

        '            'delivery_name   delivery_company    delivery_city   delivery_street_address ¡ÒÈËÏ¸Ú default=null    ºÒÔˆÔÚ default=0    delivery_postcode   ÀÔÈ‹ default=null  customers_telephone customers_fax   ’ÔÍ·Ù‹ÛÙÁÏ· default=ÃÔı„Ô˝‰ÁÚ …. & ”…¡ œ≈ –·Ò·ÙÁÒﬁÛÂÈÚ default= ÍÂÌ¸  ◊Ò›˘ÛÁ default=A    ‘ÂÏ‹˜È· default=1   ¬‹ÒÔÚ default = 2   if payment_method=¡ÌÙÈÍ·Ù·‚ÔÎﬁ = value  ‘Ò¸ÔÚ ÎÁÒ˘ÏﬁÚ default =M  ¡Ûˆ‹ÎÂÈ· default =null   ÂÌÙ default=null   ”˜ÂÙÈÍ¸ 1 default=orders_id ”˜ÂÙÈÍ¸ 2 default=null  øÒ· –·Ò‹‰ÔÛÁÚ default=null  –ÒÔﬂÔÌÙ· default=null















        '244:        Adodc1.Recordset.MoveNext()
        '246:        'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            rr = rr + 1

        '        Loop
        '248:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.SaveAs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.SaveAs("c:\ACS.xlsx")


        '250:    Excel.Quit()
        '        Exit Sub
        '        '=========================================================================
    End Sub



    Sub GENIKI()
        '        Dim antik As Integer
        '        Dim k, rr As Object


        '        Dim Excel As Microsoft.Office.Interop.Excel.Application
        '        'UPGRADE_ISSUE: Excel.workbook object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        Dim workbook As Excel.Workbook
        '        'UPGRADE_ISSUE: Excel.Worksheet object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        Dim myXL As Excel.Worksheet
        '100:    Excel = CreateObject("excel.Application")
        '102:    workbook = ExcelGlobal_definst.Workbooks.Add




        '        '¡Õ¡√ÕŸ—…”‘… œ –≈À¡‘«     ≈Õ‘—œ  œ”‘œ’”     œÕœÃ¡           ƒ…≈’»’Õ”«                –œÀ«              ‘«À≈÷ŸÕœ                        ‘                      –—œœ—…”Ãœ”    œÕœÃ¡  œ’—…≈—                ‘≈Ã¡◊…¡               ¬¡—œ”                 –¡—¡‘«—«”≈…”                 –—œ”»≈‘≈” ’–«—≈”…≈” –œ”œ ¡Õ‘… ¡‘¡¬œÀ«”  –œ”œ ¡”÷¡À…”«”
        '        'orders_id                 default=null    delivery_name   delivery_street_address delivery_city   customers_telephone & customers_fax delivery_postcode   delivery_suburb  ºÌÔÏ·  Ô˝ÒÈÂÒ default=null  ‘ÂÏ‹˜È· default=1   ¬‹ÒÔÚ default = 2   –·Ò·ÙÁÒﬁÛÂÈÚ default= ÍÂÌ¸  –Ò¸ÛËÂÙÂÚ ’ÁÒÂÛﬂÂÚ if payment_method=¡ÌÙÈÍ·Ù·‚ÔÎﬁ then =¡Õ‘… ¡‘¡¬œÀ« Ã≈‘—«‘œ…” if payment_method=¡ÌÙÈÍ·Ù·‚ÔÎﬁ = value  –¸ÛÔ default =null
        '        'On Error GoTo Command1_Click_Err
        '252:    'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        rr = 1
        '254:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 1) = "¡Õ¡√ÕŸ—…”‘… œ –≈À¡‘«"
        '256:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 2) = " ≈Õ‘—œ  œ”‘œ’”"
        '258:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 3) = "œÕœÃ¡"
        '260:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 4) = "ƒ…≈’»’Õ”«"
        '262:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 5) = "–œÀ«"
        '264:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 6) = "‘«À≈÷ŸÕœ"
        '266:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 7) = "‘ "
        '268:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 8) = "–—œœ—…”Ãœ”"
        '270:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 9) = "œÕœÃ¡  œ’—…≈—"
        '272:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 10) = "‘≈Ã¡◊…¡"
        '274:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 11) = "¬¡—œ”"
        '276:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 12) = "–¡—¡‘«—«”≈…”"
        '278:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 13) = "–—œ”»≈‘≈” ’–«—≈”…≈”"

        '280:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 14) = "–œ”œ ¡Õ‘… ¡‘¡¬œÀ«”"
        '282:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 15) = "–œ”œ ¡”÷¡À…”«”"


        '284:    'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        rr = 2
        '286:    Do While Not Adodc1.Recordset.EOF

        '288:        ANTIK = 0
        '290:        If VB.Left(Adodc1.Recordset.Fields("payment_method").Value, 5) = "¡ÌÙÈÍ" Then

        '292:            fr2.Open("select * from orders_total where orders_id=" & Str(Adodc1.Recordset.Fields("orders_id").Value) & " and class ='ot_total'", GdbS, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        '                '   + " and class IN ('ot_fixed_payment_chg','ot_shipping')"

        '294:            If Not fr2.EOF Then
        '296:                ANTIK = System.Math.Round(fr2.Fields("value").Value, 2)

        '                    'Do While Not fr2.EOF
        '                    '  For k = 0 To fr2.Fields.Count - 1
        '                    '     myXL.cells(rr, k + 1) = fr2(k).Name
        '                    '     myXL.cells(rr + 1, k + 1) = fr2(k)
        '                    '  Next
        '                    '  rr = rr + 2
        '                    '  fr2.MoveNext
        '                    'Loop
        '                End If
        '298:            fr2.Close()
        '            End If

        '            ' GENIKI-------------------------------------------

        '300:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 1) = Adodc1.Recordset.Fields("orders_id").Value
        '            'myXL.cells(rr, 2) = Adodc1.Recordset("delivery_company")
        '302:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 3) = Adodc1.Recordset.Fields("delivery_name").Value
        '304:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 4) = Adodc1.Recordset.Fields("delivery_street_address").Value
        '306:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 5) = Adodc1.Recordset.Fields("delivery_city").Value
        '308:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 6) = Adodc1.Recordset.Fields("customers_telephone").Value

        '310:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 7) = Adodc1.Recordset.Fields("delivery_postcode").Value
        '312:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 8) = Adodc1.Recordset.Fields("delivery_suburb").Value

        '            ' myXL.cells(rr, 9) = Adodc1.Recordset("customers_telephone")
        '314:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 10) = 1 'TEM
        '316:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 11) = 2 'BAROS '
        '318:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 12) = " "
        '320:        If ANTIK > 0 Then
        '322:            'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                myXL.cells(rr, 13) = "¡Õ‘… ¡‘¡¬œÀ« Ã≈‘—«‘œ…”"
        '            Else
        '324:            'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                myXL.cells(rr, 13) = "" ' "◊—≈Ÿ”«"
        '            End If
        '326:        If ANTIK > 0 Then
        '328:            'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                myXL.cells(rr, 14) = VB6.Format(System.Math.Round(ANTIK, 2), "#####.##")
        '            End If


        '330:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 15) = "" ' "‘—œ–œ” –À«—ŸÃ«”"

        '332:        Adodc1.Recordset.MoveNext()
        '334:        'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            rr = rr + 1

        '        Loop
        '336:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.SaveAs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.SaveAs("c:\geniki.xlsx")


        '338:    Excel.Quit()
        '        Exit Sub
        '        '===============================================================

    End Sub

    Sub PORTA()
        '        'porta
        '        Dim antik As Integer
        '        Dim k, rr As Object


        '        Dim Excel As Microsoft.Office.Interop.Excel.Application
        '        'UPGRADE_ISSUE: Excel.workbook object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        Dim workbook As Excel.Workbook
        '        'UPGRADE_ISSUE: Excel.Worksheet object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        Dim myXL As Excel.Worksheet
        '100:    Excel = CreateObject("excel.Application")
        '102:    workbook = ExcelGlobal_definst.Workbooks.Add




        '        '1  delivery_name
        '        '2  delivery_street_address
        '        '3  delivery_city   '
        '        '4  delivery_postcode
        '        '5  customers_telephone
        '        '6  customers_fax   '
        '        '7  ‘ÂÏ‹˜È· default=1   '
        '        '8  ¬‹ÒÔÚ default = 3
        '        '9  ”˜¸ÎÈ· default= ÍÂÌ¸
        '        '10 if payment_method=¡ÌÙÈÍ·Ù·‚ÔÎﬁ = value
        '        '11 ≈–…‘¡√« default=null
        '        '12 À«Œ« ≈–…‘¡√«” default=null
        '        '13 ƒÁÎ˘Ï›ÌÁ ¡Óﬂ· default=null
        '        '14 orders_id
        '        '15 ≈È‚‹ÒıÌÛÁ 1 default=null
        '        '16 ≈È‚‹ÒıÌÛÁ 2 default=null
        '        '17 ≈È‚‹ÒıÌÛÁ 3 default=null
        '340:    'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        rr = 1
        '342:    Do While Not Adodc1.Recordset.EOF
        '344:        ANTIK = 0
        '346:        If VB.Left(Adodc1.Recordset.Fields("payment_method").Value, 5) = "¡ÌÙÈÍ" Then
        '348:            'fr2.open "select * from orders_total where orders_id=" + Str(Adodc1.Recordset("orders_id")) + " and class ='ot_total'", GdbS, adOpenDynamic, adLockOptimistic
        '350:            ' If Not fr2.EOF Then
        '352:            ANTIK = System.Math.Round(Adodc1.Recordset.Fields("total").Value, 2)
        '            End If
        '354:        fr2.Close()

        '356:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 1) = Adodc1.Recordset.Fields("shipping_firstname").Value + " " + Adodc1.Recordset.Fields("shipping_Lastname").Value
        '358:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 2) = Adodc1.Recordset.Fields("shipping_address_1").Value
        '360:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 3) = Adodc1.Recordset.Fields("shipping_city").Value
        '362:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 4) = Adodc1.Recordset.Fields("shipping_postcode").Value
        '364:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 5) = Adodc1.Recordset.Fields("telephone").Value
        '366:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 6) = Adodc1.Recordset.Fields("telephone").Value

        '368:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 7) = 1 ' Adodc1.Recordset("delivery_postcode")
        '370:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 8) = 1 '  "Àœ…–¡"
        '            ' myXL.cells(rr, 9) = Adodc1.Recordset("customers_telephone")


        '372:        If ANTIK > 0 Then
        '374:            'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                myXL.cells(rr, 10) = VB6.Format(System.Math.Round(ANTIK, 2), "#####.##")
        '            End If

        '            'myXL.cells(rr, 11) = "ÃÔı„Ô˝‰ÁÚ …. & ”…¡ œ≈"

        '            'myXL.cells(rr, 12) = " "
        '376:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 13) = "" ' "◊—≈Ÿ”«"
        '378:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 14) = Adodc1.Recordset.Fields("order_id").Value

        '            'myXL.cells(rr, 15) = 2 'BAROS
        '            ' myXL.cells(rr, 16) = Format(Round(ANTIK, 2), "#####.##")
        '            ' myXL.cells(rr, 17) = "M"  ' "‘—œ–œ” –À«—ŸÃ«”"

        '380:        Adodc1.Recordset.MoveNext()
        '382:        'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            rr = rr + 1

        '        Loop
        '384:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.SaveAs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.SaveAs("c:\PORTA.xlsx")


        '386:    Excel.Quit()

        '        '<EhFooter>
        '        Exit Sub
    End Sub

    '==========================================================================
    Sub speedex()  ' ARXH
        '        Dim antik As Integer
        '        Dim k, rr As Object


        '        Dim Excel As Microsoft.Office.Interop.Excel.Application
        '        'UPGRADE_ISSUE: Excel.workbook object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        Dim workbook As Excel.Workbook
        '        'UPGRADE_ISSUE: Excel.Worksheet object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        Dim myXL As Excel.Worksheet
        '100:    Excel = CreateObject("excel.Application")
        '102:    workbook = ExcelGlobal_definst.Workbooks.Add


        '        '==============================================================================
        '        'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        rr = 1
        '        '146   myXL.cells(rr, 1) = "œÕœÃ¡ –¡—¡À«–‘«"
        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 1) = "≈–ŸÕ’Ã…¡ ≈‘¡…—…¡”"
        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 2) = "ƒ…≈’»’Õ”« –¡—¡À«–‘«"


        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 3) = "–≈—…œ◊«"
        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 4) = "–œÀ«"
        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 5) = "‘ "
        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 6) = "‘«À≈÷ŸÕœ"
        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 7) = "¡Õ‘… "
        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 8) = "–œ”œ ¡Õ‘… "
        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 9) = "œÀœ√—¡÷Ÿ”"
        '        'myXL.cells(rr, 10) = ""
        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 11) = ""
        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 10) = "–¡—¡‘«—«”≈…”"
        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.cells(rr, 12) = ""



        '        'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        rr = 2
        '1342:   Do While Not Adodc1.Recordset.EOF
        '1344:       ANTIK = 0
        '1346:       If VB.Left(Adodc1.Recordset.Fields("payment_method").Value, 5) = "¡ÌÙÈÍ" Then
        '1348:           fr2.Open("select * from orders_total where orders_id=" & Str(Adodc1.Recordset.Fields("orders_id").Value) & " and class ='ot_total'", GdbS, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        '1350:           If Not fr2.EOF Then
        '1352:               ANTIK = System.Math.Round(fr2.Fields("value").Value, 2)
        '                End If
        '1354:           fr2.Close()
        '            End If
        '1356:       'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 1) = Adodc1.Recordset.Fields("delivery_name").Value
        '1358:       'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 2) = Adodc1.Recordset.Fields("delivery_street_address").Value
        '1360:       'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 3) = Adodc1.Recordset.Fields("delivery_city").Value
        '            'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 4) = Adodc1.Recordset.Fields("delivery_city").Value

        '1362:       'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 5) = Adodc1.Recordset.Fields("delivery_postcode").Value

        '1364:       'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 6) = Adodc1.Recordset.Fields("customers_telephone").Value


        '            If ANTIK > 0 Then
        '                'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                myXL.cells(rr, 7) = 1
        '                'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                myXL.cells(rr, 8) = VB6.Format(System.Math.Round(ANTIK, 2), "#####.##")
        '            Else
        '                'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                myXL.cells(rr, 7) = 0
        '                'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                myXL.cells(rr, 8) = 0 ' Format(Round(ANTIK, 2), "#####.##")
        '            End If

        '            'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            myXL.cells(rr, 9) = "M" ' Adodc1.Recordset("customers_fax")


        '            '
        '            '
        '            '
        '            '
        '            '368   myXL.cells(rr, 7) = 1 ' Adodc1.Recordset("delivery_postcode")
        '            '370   myXL.cells(rr, 8) = 2 '  "Àœ…–¡"
        '            '     ' myXL.cells(rr, 9) = Adodc1.Recordset("customers_telephone")
        '            '
        '            '
        '            '   If ANTIK > 0 Then
        '            '      myXL.cells(rr, 10) = Format(Round(ANTIK, 2), "#####.##")
        '            '   End If
        '            '
        '            '      'myXL.cells(rr, 11) = "ÃÔı„Ô˝‰ÁÚ …. & ”…¡ œ≈"
        '            '
        '            '      'myXL.cells(rr, 12) = " "
        '            '376   myXL.cells(rr, 13) = ""   ' "◊—≈Ÿ”«"
        '            '378   myXL.cells(rr, 14) = Adodc1.Recordset("orders_id")

        '1380:       Adodc1.Recordset.MoveNext()
        '1382:       'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            rr = rr + 1

        '        Loop

        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Columns("A:K").Select()
        '        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.Columns.AutoFit()



        '1384:   'UPGRADE_WARNING: Couldn't resolve default property of object myXL.SaveAs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.SaveAs("c:\EASY PRINTNEW.xlsx")
        '        ' myXL.Activate

        '        Excel.Visible = True


        '        Dim ANS3 As Integer

        '        ANS3 = MsgBox(" ÎÂﬂÌ˘ ÙÔ EXCEL", MsgBoxStyle.YesNo)

        '        If ANS3 = MsgBoxResult.Yes Then
        '            'UPGRADE_WARNING: Couldn't resolve default property of object workbook.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            Call workbook.Close(False)
        '            Excel.Quit()
        '            'UPGRADE_NOTE: Object Excel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '            Excel = Nothing
        '        End If


        '        '1386 Excel.Quit

        '        '<EhFooter>
        '        Exit Sub


        '        '==========================================================================
        '        '  */   speedex: TELOS
        '        '==============================================================================

    End Sub



    'Command1_Click_Err:
    '       SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.Command1_Click " & "at line " & Erl())
    '      Resume Next
    '</EhFooter>
    '  End Sub

    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click


        '        '<EhHeader>
        '        On Error GoTo Command2_Click_Err
        '        '</EhHeader>
        '        Dim Excel As Microsoft.Office.Interop.Excel.Application
        '        'UPGRADE_ISSUE: Excel.workbook object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        Dim workbook As Excel.Workbook
        '        'UPGRADE_ISSUE: Excel.Worksheet object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        Dim myXL As Excel.Worksheet
        '100:    Excel = CreateObject("excel.Application")
        '102:    workbook = ExcelGlobal_definst.Workbooks.Add
        '        On Error Resume Next
        '        Dim F_XROMATA As Short
        '104:    'UPGRADE_WARNING: Couldn't resolve default property of object workbook.Activate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        workbook.Activate()

        '106:    'UPGRADE_WARNING: Couldn't resolve default property of object workbook.ActiveSheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL = workbook.ActiveSheet




        '        Dim k, rr As Object




        '108:    Adodc1.Recordset.MoveFirst()

        '110:    'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        rr = 1

        '        Dim AA As Short
        '        ' AA = InputBox("1.ACS  2.√≈Õ… « 3.–œ—‘¡-–œ—‘¡ 4.SPEEDEX", , 1)
        '        AA = 3

        '        '¡CS
        '        'œÕœÃ¡ –¡—¡À«–‘« ≈–ŸÕ’Ã…¡ ≈‘¡…—…¡”   –≈—…œ◊« œƒœ”    ¡—…»Ãœ” œ—œ÷œ”  ‘   Àœ…–¡   ‘«À≈÷ŸÕœ     …Õ«‘œ  ’–œ ¡‘¡”‘«Ã¡    –¡—¡‘«—«”≈…”    ◊—≈Ÿ”«  ‘≈Ã¡◊…¡ ¬¡—œ”   ¡Õ‘… ¡‘¡¬œÀ«    ‘—œ–œ” –À«—ŸÃ«” ¡”÷¡À≈…¡     ≈Õ‘—œ  œ”‘œ’”  ”◊≈‘… œ1    ”◊≈‘… œ2    Ÿ—¡ –¡—¡ƒœ”«”   –—œ…œÕ‘¡
        '        'delivery_name   delivery_company    delivery_city   delivery_street_address ¡ÒÈËÏ¸Ú default=null    ºÒÔˆÔÚ default=0    delivery_postcode   ÀÔÈ‹ default=null  customers_telephone customers_fax   ’ÔÍ·Ù‹ÛÙÁÏ· default=ÃÔı„Ô˝‰ÁÚ …. & ”…¡ œ≈ –·Ò·ÙÁÒﬁÛÂÈÚ default= ÍÂÌ¸  ◊Ò›˘ÛÁ default=A    ‘ÂÏ‹˜È· default=1   ¬‹ÒÔÚ default = 2   if payment_method=¡ÌÙÈÍ·Ù·‚ÔÎﬁ = value  ‘Ò¸ÔÚ ÎÁÒ˘ÏﬁÚ default =M  ¡Ûˆ‹ÎÂÈ· default =null   ÂÌÙ default=null   ”˜ÂÙÈÍ¸ 1 default=orders_id ”˜ÂÙÈÍ¸ 2 default=null  øÒ· –·Ò‹‰ÔÛÁÚ default=null  –ÒÔﬂÔÌÙ· default=null
        '        Dim ANTIK As Single

        '120:    Do While Not Adodc1.Recordset.EOF

        '122:        ANTIK = 0
        '124:        fr2.Open("select * from orders_total where order_status_id=" & Str(Adodc1.Recordset.Fields("orders_id").Value) & " and class='ot_fixed_payment_chg'")
        '126:        If Not fr2.EOF Then
        '128:            ANTIK = fr2.Fields("VALUE").Value
        '            End If
        '130:        fr2.Close()

        '132:        For k = 0 To Adodc1.Recordset.Fields.Count - 1
        '134:            'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                myXL.Cells(rr, k + 1) = Adodc1.Recordset.Fields(k).Value
        '            Next
        '136:        Adodc1.Recordset.MoveNext()
        '138:        'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            rr = rr + 1

        '        Loop


        '140:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.SaveAs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        myXL.SaveAs("c:\parag.xlsx")


        '142:    Excel.Quit()




        'PORTA:
        '        'porta

        '        '1  delivery_name
        '        '2  delivery_street_address
        '        '3  delivery_city   '
        '        '4  delivery_postcode
        '        '5  customers_telephone
        '        '6  customers_fax   '
        '        '7  ‘ÂÏ‹˜È· default=1   '
        '        '8  ¬‹ÒÔÚ default = 3
        '        '9  ”˜¸ÎÈ· default= ÍÂÌ¸
        '        '10 if payment_method=¡ÌÙÈÍ·Ù·‚ÔÎﬁ = value
        '        '11 ≈–…‘¡√« default=null
        '        '12 À«Œ« ≈–…‘¡√«” default=null
        '        '13 ƒÁÎ˘Ï›ÌÁ ¡Óﬂ· default=null
        '        '14 orders_id
        '        '15 ≈È‚‹ÒıÌÛÁ 1 default=null
        '        '16 ≈È‚‹ÒıÌÛÁ 2 default=null
        '        '17 ≈È‚‹ÒıÌÛÁ 3 default=null
        '340:    'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        rr = 1

        '        '        Adodc1.Recordset.MoveFirst()

        '        '342:    Do While Not Adodc1.Recordset.EOF
        '        '344:        ANTIK = 0
        '        '346:        If VB.Left(Adodc1.Recordset.Fields("payment_method").Value, 5) = "¡ÌÙÈÍ" Then
        '        '348:            'fr2.open "select * from orders_total where orders_id=" + Str(Adodc1.Recordset("orders_id")) + " and class ='ot_total'", GdbS, adOpenDynamic, adLockOptimistic
        '        '350:            ' If Not fr2.EOF Then
        '        '352:            ANTIK = System.Math.Round(Adodc1.Recordset.Fields("total").Value, 2)
        '        '            End If
        '        '354:        fr2.Close()

        '        '356:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '            myXL.Cells(rr, 1) = Adodc1.Recordset.Fields("shipping_firstname").Value + " " + Adodc1.Recordset.Fields("shipping_Lastname").Value
        '        '358:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '            myXL.Cells(rr, 2) = Adodc1.Recordset.Fields("shipping_address_1").Value
        '        '360:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '            myXL.Cells(rr, 3) = Adodc1.Recordset.Fields("shipping_city").Value
        '        '362:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '            myXL.Cells(rr, 4) = Adodc1.Recordset.Fields("shipping_postcode").Value
        '        '364:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '            myXL.Cells(rr, 5) = Adodc1.Recordset.Fields("telephone").Value
        '        '366:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '            myXL.Cells(rr, 6) = Adodc1.Recordset.Fields("telephone").Value

        '        '368:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '            myXL.Cells(rr, 7) = 1 ' Adodc1.Recordset("delivery_postcode")
        '        '370:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '            myXL.Cells(rr, 8) = 2 '  "Àœ…–¡"
        '        '            ' myXL.cells(rr, 9) = Adodc1.Recordset("customers_telephone")


        '        '372:        If ANTIK > 0 Then
        '        '374:            'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '                myXL.Cells(rr, 10) = VB6.Format(System.Math.Round(ANTIK, 2), "#####.##")
        '        '            End If

        '        '            'myXL.cells(rr, 11) = "ÃÔı„Ô˝‰ÁÚ …. & ”…¡ œ≈"

        '        '            'myXL.cells(rr, 12) = " "
        '        '376:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '            myXL.Cells(rr, 13) = "" ' "◊—≈Ÿ”«"
        '        '378:        'UPGRADE_WARNING: Couldn't resolve default property of object myXL.cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '            myXL.Cells(rr, 14) = Adodc1.Recordset.Fields("order_id").Value

        '        '            'myXL.cells(rr, 15) = 2 'BAROS
        '        '            ' myXL.cells(rr, 16) = Format(Round(ANTIK, 2), "#####.##")
        '        '            ' myXL.cells(rr, 17) = "M"  ' "‘—œ–œ” –À«—ŸÃ«”"

        '        '380:        Adodc1.Recordset.MoveNext()
        '        '382:        'UPGRADE_WARNING: Couldn't resolve default property of object rr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '            rr = rr + 1

        '        '        Loop
        '        '384:    'UPGRADE_WARNING: Couldn't resolve default property of object myXL.SaveAs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        '        myXL.SaveAs("c:\PORTA.xlsx")


        '        '386:    Excel.Quit()

        '        '<EhFooter>
        '        Exit Sub





        'Command2_Click_Err:
        '        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.Command2_Click " & "at line " & Erl())
        '        Resume Next
        '        '</EhFooter>



















    End Sub



    Private Sub DataGrid1_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        '  If InStr(Text1.Text, DataGrid1.Columns(0).Text) > 0 Then
        '     'Me.Text = "ı·Ò˜ÂÈ Á‰Á"
        '    Else
        '    Text1.Text = Text1.Text & DataGrid1.Columns(0).Text & ","
        '    Me.Text = ""
        '    End If
    End Sub

    Private Sub EXEC_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EXEC.Click



    End Sub


    Private Function UnicodeStringToBytes(ByVal str As String) As Byte()

        Return System.Text.Encoding.Unicode.GetBytes(str)
    End Function
    Private Sub update_order(ByVal mRow As Integer)
        ' Dim k As Object
        Dim n As Object
        Dim AJ5 As Object
        Dim AJ3 As Object
        Dim AJ7 As Object
        Dim AJ6 As Object
        Dim EPA As Object
        '<EhHeader>
        On Error GoTo EXEC_Click_Err
        '</EhHeader>
        Dim poso, kod, timm As Object
        Dim nk As Object
        Dim onoma As Object
        Dim fpa As Object
        Dim DD(300) As Object
        Dim POL, EPO, DIE, PERIOXH As Object
        Dim SQL As String
        Dim PELKOD As Object
        Dim THL As Object
        Dim afm As String
        Dim doy As String



        'customes_?,billing_?,delivery_?
        'name,company,street_address,suburb,city,postcode,state,country,address_format_id




        '        √È˛Ò„Ô Í·ÎÁÛ›Ò·,

        '‰Â ÛÔı Ù· Âﬂ· ÛÙÔ ÙÁÎ›ˆ˘ÌÔ ·ÍÒÈ‚˛Ú ¸˘Ú ÂﬂÌ·È ‰È¸ÙÈ ÏÂÒ‰Â˝ÙÁÍ·, Âﬂ˜· Í‹ÙÈ ‹ÎÎÔ ÛÙÔ Ïı·Î¸ ÏÔı.

        '‘Ô opencart Ôı ˜ÒÁÛÈÏÔÔÈÂﬂ ÙÔ youhou.gr ›˜ÂÈ ÙÂÎÈÍ‹ flag „È· ¡¸‰ÂÈÓÁ ﬁ ‘ÈÏÔÎ¸„ÈÔ ÛÙÔÌ ﬂÌ·Í· oc_order Í·È ÂﬂÌ·È Á ÛÙﬁÎÁ customer_group_id. 

        '¡Ì ›˜ÂÈ ÙÈÏﬁ 1 ÂﬂÌ·È ¡¸‰ÂÈÓÁ ÂÌ˛ ÙÈÏﬁ 2 ÂﬂÌ·È ‘ÈÏÔÎ¸„ÈÔ

        '”ÙÔÌ ﬂ‰ÈÔ ﬂÌ·Í· oc_order, ÛÙÁ ÛÙﬁÎÁ custom_field, ¸Ù·Ì ›˜ÔıÏÂ ÙÈÏÔÎ¸„ÈÔ ı‹Ò˜ÂÈ ÏÈ· ÛıÏ‚ÔÎÔÛÂÈÒ‹ ÊÂı„˛Ì ÙÈÏ˛Ì (ﬂÌ·Í·Ú  JSON) ¸Ôı ‚ÒﬂÛÍÔÌÙ·È Ù· ÛÙÔÈ˜Âﬂ· Í˘‰ÈÍÔÔÈÁÏ›Ì· ÛÂ UNICODE :

        '        –…Õ¡ ¡”(1)
        '{"1":"\u03b5\u03bc\u03c0\u03bf\u03c1\u03b9\u03ba\u03b7 \u03b1\u03bd\u03c9\u03bd\u03c5\u03bc\u03b7 ","2":"095466356","3":"\u03a3\u0391\u039c\u039f\u03a5"}

        '‘Ô ·Ò·‹Ì˘ Ï·Ú ‰ﬂÌÂÈ :

        '{"1":"ÂÏÔÒÈÍÁ ·Ì˘ÌıÏÁ ","2":"095466356","3":"”¡Ãœ’"}

        '”ÙÔ 1 Ò›ÂÈ Ì· ›˜ÔıÏÂ ÙÔ ≈‹„„ÂÎÏ·, ÛÙÔ 2 ÙÔ ¡÷Ã Í·È ÛÙÔ 3 ÙÁ ƒœ’

        If SQLDT2.Rows(mRow).Item("customer_group_id") = 2 Then

            Dim ffff As String = utftoascii(SQLDT2.Rows(mRow).Item("custom_field").ToString)

            'Dim cs() As Byte = System.Text.Encoding.Default.GetBytes(SQLDT2.Rows(mRow).Item("custom_field").ToString)

            'Dim cs2() As Byte = UnicodeStringToBytes(SQLDT2.Rows(mRow).Item("custom_field").ToString)
            'Dim ff As String = ""
            'For ffk As Integer = 1 To 200   'Length(cs)  
            '    If cs(ffk) > 32 Then
            '        ff = ff + Chr(cs(ffk))
            '        Debug.Print(ff)
            '    End If

            'Next


            ' Dim cc2 As String = UnicodeBytesToString(SQLDT2.Rows(mRow).Item("customer_group_id").ToString)


            ffff = Split(ffff, ":")(2).ToString
            afm = Split(ffff, ",")(0).ToString

            afm = Replace(afm, """", "")

        Else
            afm = ""

        End If



        'If Len(DD(6)) = 0 Then ' …ƒ…Ÿ‘«”
100:    'UPGRADE_WARNING: Couldn't resolve default property of object EPO. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        EPO = SQLDT2.Rows(mRow).Item("payment_lastname").ToString + " " + SQLDT2.Rows(mRow).Item("payment_FIRSTname").ToString
        'Else ' ETAIREIA +1 SEIRA

        'End If


        Dim CH1 As String
        Dim EMAIL As String

        Dim TK As Object
102:    'UPGRADE_WARNING: Couldn't resolve default property of object DIE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        DIE = Replace(Mid(SQLDT2.Rows(mRow).Item("payment_address_1").ToString, 1, 35), "'", "`")


        EMAIL = Replace(Mid(SQLDT2.Rows(mRow).Item("email").ToString, 1, 35), "'", "`")
        CH1 = Trim(Replace(Mid(SQLDT2.Rows(mRow).Item("comment").ToString, 1, 120), "'", "`"))

104:    'UPGRADE_WARNING: Couldn't resolve default property of object POL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        POL = Replace(Mid(SQLDT2.Rows(mRow).Item("payment_city").ToString, 1, 35), "'", "`")

106:    'UPGRADE_WARNING: Couldn't resolve default property of object EPA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

        If IsDBNull(SQLDT2.Rows(mRow).Item("epaggelma")) Then
            EPA = "…ƒ…Ÿ‘«”"
        Else
            EPA = SQLDT2.Rows(mRow).Item("epaggelma").ToString
        End If

        'If IsDBNull(SQLDT2.Rows(mRow).Item("afm")) Then
        '    afm = " "
        'Else
        '    afm = SQLDT2.Rows(mRow).Item("afm").ToString
        'End If


        If IsDBNull(SQLDT2.Rows(mRow).Item("doy")) Then
            doy = " "
        Else
            doy = SQLDT2.Rows(mRow).Item("doy").ToString
        End If





        ' EPA = " " ' Replace(Mid(Adodc1.Recordset("payment_suburb"), 1, 35), "'", "`")
        'UPGRADE_WARNING: Couldn't resolve default property of object TK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TK = Replace(Mid(SQLDT2.Rows(mRow).Item("payment_POSTCODE").ToString, 1, 35), "'", "`")
        'If InStr(DD(TSONTA + 5), "<") - 1 < 1 Then
        '   ch1 = ""
        'Else
        '   ch1 = Mid(DD(TSONTA + 5), 1, InStr(DD(TSONTA + 5), "<") - 1)
        'End If
108:    'UPGRADE_WARNING: Couldn't resolve default property of object PELKOD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PELKOD = "C" & VB6.Format(SQLDT2.Rows(mRow).Item("order_id").ToString, "000000")

110:    'UPGRADE_WARNING: Couldn't resolve default property of object THL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        THL = SQLDT2.Rows(mRow).Item("telephone").ToString

        Dim noTEL As Short
        noTEL = 1
        '   If Len(THL) < 6 Then
        '        noTEL = 1
        '   End If


        FileOpen(2, "c:\mercvb\errweb.txt", OpenMode.Append)






        '        PrintLine(2, A)












        Dim KPE As String '  œƒ… œ” –≈À¡‘«
        Dim r As New ADODB.Recordset
112:    'UPGRADE_WARNING: Couldn't resolve default property of object THL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        r.Open("select * from PPEL where EIDOS='e' and THL='" + THL + "'", Gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
114:    If r.EOF Or noTEL = 1 Then
116:        r.Close()
118:        'UPGRADE_WARNING: Couldn't resolve default property of object PELKOD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            r.Open("select count(*) from PPEL where EIDOS='e' and KOD='" + PELKOD + "'", Gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
120:        If r.Fields(0).Value = 0 Then
122:            SQL = "INSERT INTO PPEL (EIDOS,KOD,EPO,DIE,POL,EPA,MEMO,THL,XRVMA,EMAIL,AFM,DOY) VALUES("
124:
                SQL = SQL & " 'e','" + PELKOD + "','" + EPO + "','" + DIE + "','" + VB.Left(POL, 10) + "','" + EPA + "','" + CH1 + "','" + THL + "','" + TK + "','" + EMAIL + "','" + afm + "','" + doy + "')"
126:            Gdb.Execute(SQL)
                PrintLine(2, SQL)
            End If
128:
            KPE = PELKOD
130:        r.Close()
        Else
132:        KPE = r.Fields("KOD").Value
134:        r.Close()

        End If

        Dim aj4, aj1, aj2, aji As Object
136:
        aj1 = 0

        aj2 = 0

        aj4 = 0

        aji = 0

        AJ6 = 0

        Dim SHMERA As String
138:    SHMERA = VB6.Format(Now, "MM/DD/YYYY")



140:
        r.Open("select count(*) from PTIM where  ATIM='a" & Mid(PELKOD, 2, 6) & "' AND HME='" & SHMERA & "' ", Gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
142:    If r.Fields(0).Value > 0 Then
144:
            Gdb.Execute("DELETE from PTIM where  ATIM='a" & Mid(PELKOD, 2, 6) & "' AND HME='" & SHMERA & "'")
146:
            Gdb.Execute("DELETE from PPEGGTIM where  ATIM='a" & Mid(PELKOD, 2, 6) & "' AND HME='" & SHMERA & "'")
        End If

148:    r.Close()


        '"orders_products_id";"orders_id";"products_id";"products_model";"products_name";"products_price";"final_price";"products_tax";"products_quantity"
        '40;14;742;"5204275010251";"";"33.61";"33.61";"19.00";1
        '37;12;88;"400878904270";"";"84.03";"84.03";"19.00";1


        'fr2.Close

        Dim pos01, pos03 As Single


        Dim RT As New ADODB.Recordset
        Dim mID_NUM As Integer



        RT.Open("SELECT MAX(ID_NUM) FROM PTIM", Gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        mID_NUM = RT.Fields(0).Value + 1

        RT.Close()

        'SELECT `oc_order_product`.`order_product_id`,
        '    `oc_order_product`.`order_id`,
        '    `oc_order_product`.`product_id`,
        '    `oc_order_product`.`name`,
        '    `oc_order_product`.`model`,
        '    `oc_order_product`.`quantity`,
        '    `oc_order_product`.`price`,
        '    `oc_order_product`.`total`,
        '    `oc_order_product`.`tax`,
        '    `oc_order_product`.`reward`
        'FROM `web88_youdb`.`oc_order_product`;



        AJ7 = 0

        aj1 = 0

        aj2 = 0

        AJ3 = 0

        aj4 = 0

        AJ5 = 0

        AJ6 = 0

        Dim k As Integer
        Dim fr2data As New DataTable
        Dim fr3data As New DataTable
        ExecuteSQLQuery("select * from oc_order_product  where order_id =" & LTrim((SQLDT2.Rows(mRow).Item("order_id").ToString)), fr2data)


        ' fr2.Open("select * from oc_order_product  where order_id =" & LTrim(Str(Adodc1.Recordset.Fields("order_id").Value)), GdbS, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
152:    Dim m_kodeid As String
        For n = 0 To fr2data.Rows.Count - 1
            'brhke seira paraggelias
154:        If Val(fr2data.Rows(n).Item("quantity").ToString) > 0 Then

                If Len(Trim(fr2data.Rows(n).Item("model").ToString)) = 0 Then
                    MsgBox("–—œ”œ◊« ƒ≈Õ ≈◊≈…  Ÿƒ… œ ‘œ " + fr2data.Rows(n).Item("Name").Value)
                    End
                End If










                If Len(fr2data.Rows(n).Item("model").ToString) > 12 Then
                    m_kodeid = "ID" & VB6.Format(fr2data.Rows(n).Item("PRODUCT_ID").ToString, "00000")

                Else
                    m_kodeid = fr2data.Rows(n).Item("model").ToString
                End If









                '156:            r.Open("select * from EID  WHERE KOD='" & m_kodeid & "'", Gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                '158:            If r.EOF Then
                '160:                'Gdb.Execute "insert into EID (KOD,ONO,FPA) VALUES('" + m_kodeid + "','" + Replace(fr2("name"), "'", "`") + "',6)", n
                '162:                'Gdb.Execute "insert into EID (KOD,ONO,FPA) VALUES('" + m_kodeid + "','" + Left(Replace(fr2("name"), 35), "'", "`") + "',6)", n
                '                    'Gdb.Execute("insert into EID (KOD,ONO,FPA) VALUES('" & m_kodeid & "','" & VB.Left(Replace(fr2.Fields("name").ToString, "'", "`"), 35) & "',6)", n)
                '                    r.Close()
                '164:                r.Open("select * from EID  WHERE KOD='" & m_kodeid & "'", Gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                '                End If


                '166:
                '                poso = fr2data.Rows(n).Item("quantity").Value

                '168:
                '                pos01 = IIf(IsDBNull(r.Fields("pos01").Value), 0, r.Fields("pos01").Value)
                '170:
                '                pos03 = IIf(IsDBNull(r.Fields("pos03").Value), 0, r.Fields("pos03").Value)


                '172:
                '                k = k + 1
                '174:            If r.EOF Then

                '176:
                '                    onoma = ""
                '                Else
                '178:
                '                    If IsDBNull(r.Fields("ONO").Value) Then
                '180:
                '                        onoma = ""
                '                    Else
                '182:
                '                        onoma = VB.Left(Replace(fr2data.Rows(n).Item("name").ToString, "'", "''"), 35) ' r!ONO
                '                    End If

                '                End If
                '184:            r.Close()

186:
                kod = m_kodeid ' fr2!products_model
188:
                fpa = fr2data.Rows(n).Item("tax").ToString

190:
                timm = fr2data.Rows(n).Item("Price").ToString

192:            SQL = ""
194:
                'onoma = VB.Left(Replace(fr2data.Rows(n).Item("name").ToString, "'", "''"), 35)
                ' onoma = VB.Left(Replace(fr2data.Rows(n).Item("name").ToString, "'", "''"), 35) : SQL = "insert " 'into PEGGTIM (APOT,ONOMA,HME,EIDOS,ATIM,PELKOD,KODE,POSO,TIMM,FPA,ID_NUM) VALUES ('1','" & Replace(VB.Left(onoma, 35), "'", "`") & "','" & SHMERA & "','e','"
196:
                onoma = VB.Left(Replace(fr2data.Rows(n).Item("name").ToString, "'", "''"), 35)
                SQL = "insert into PEGGTIM (APOT,ONOMA,HME,EIDOS,ATIM,PELKOD,KODE,POSO,TIMM,FPA,ID_NUM) VALUES ('1','" & Replace(VB.Left(onoma, 35), "'", "`") & "','" & SHMERA & "','e','"
                SQL = SQL & "a" & Mid(PELKOD, 2, 6) & "','"
198:            SQL = SQL & KPE & "','"
200:
                SQL = SQL + kod + "',"
202:

                poso = fr2data.Rows(n).Item("quantity").ToString
                SQL = SQL & Replace(VB6.Format(poso, "####.00"), ",", ".") & ","
204:
                SQL = SQL & Replace(VB6.Format(timm, "####.00"), ",", ".") & ","



                ' SQL = SQL + "6,"
206:
                If Val(fpa) = 24 Then
208:                '
210:                ' AJ6 = AJ6 + poso * timm
212:                'UPGRADE_WARNING: Couldn't resolve default property of object fpa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf Val(fpa) >= 16 And Val(fpa) < 18 Then
214:                'SQL = SQL + "7,"
216:                ' AJ7 = AJ7 + poso * timm
218:                'If Val(fpa) = 16 Then
220:                ' SQL = SQL + "1,"
222:                'aj1 = aj1 + poso * timm
                End If




                If System.Math.Round(Val(CStr(100 * fpa / timm)), 0) = 24 Then
                    SQL = SQL & "6,"
                    AJ6 = AJ6 + poso * timm


                ElseIf System.Math.Round(Val(CStr(100 * fpa / timm)), 0) = 23 Then
                    SQL = SQL & "6,"
                    AJ6 = AJ6 + poso * timm

                ElseIf System.Math.Round(Val(CStr(100 * fpa / timm)), 0) = 17 Then
                    SQL = SQL & "7,"
                    AJ7 = AJ7 + poso * timm

                Else
                    SQL = SQL + "6,"

                End If








                'If Val(fpa) = 0 Then
226:            'SQL = SQL & "5,"
                'AJ5 = AJ5 + poso * timm
                'End If
                SQL = SQL & Str(mID_NUM) & ")"
                Dim NIK As Integer
230:            Gdb.Execute(SQL, NIK)
                PrintLine(2, SQL)
                ' GdbS.Execute "UPDATE products set drama_quantity=" + Format(pos01, "####0.00") + ",xanthi_quantity=" + Format(pos03, "####0.00") + " where products_model='" + kod + "' limit 1 ", nk

                ' GdbS.Execute "UPDATE products set drama_quantity=0,xanthi_quantity=0 limit 1 ", nk

            End If
232:        'fr2.MoveNext()
        Next  'fr2data

234:    'fr2.Close()

        ExecuteSQLQuery("select * from oc_order_total where order_id=" & Str(SQLDT2.Rows(mRow).Item("order_id").ToString) & "  and code='shipping'", fr3data)

        'fr2.Open("select * from oc_order_total where order_id=" & Str(Adodc1.Recordset.Fields("order_id").Value) & "  and code='shipping'")
238:    Dim R5 As New ADODB.Recordset
        If fr3data.Rows.Count > 0 Then
            poso = 1
            timm = fr3data.Rows(0).Item("Value").ToString / 1.24
            kod = "METAF"
            onoma = "Ã≈‘¡÷œ—… ¡ ≈Œœƒ¡"



            R5.Open("select * from EID  WHERE KOD='" & m_kodeid & "'", Gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If R5.EOF Then
                Gdb.Execute("insert into EID (KOD,ONO,FPA) VALUES('" + kod + "','" + "Ã≈‘¡÷œ—… ¡ ≈Œœƒ¡" + "',6)", n)
            End If
            R5.Close()




248:        SQL = ""
            SQL = "insert into PEGGTIM (APOT,ONOMA,HME,EIDOS,ATIM,PELKOD,KODE,POSO,TIMM,FPA,ID_NUM) VALUES ('1','" + onoma + "','" + SHMERA + "','e','"
            SQL = SQL & "a" & Mid(PELKOD, 2, 6) & "','"
254:        SQL = SQL & KPE & "','"
            SQL = SQL + kod + "',"
            SQL = SQL & Replace(VB6.Format(poso, "####.00"), ",", ".") & ","
            SQL = SQL & Replace(VB6.Format(timm, "####.00"), ",", ".") & ","
262:        SQL = SQL & "6,"
            SQL = SQL & Str(mID_NUM) & ")"
            AJ6 = AJ6 + poso * timm
266:        Gdb.Execute(SQL, nk)
            PrintLine(2, SQL)
        End If




        '===================  antikatabolh ===================================================================================

        ExecuteSQLQuery("select * from oc_order_total where order_id=" & Str(SQLDT2.Rows(mRow).Item("order_id").ToString) & "  and code='GOP_COD_Fee'", fr3data)

        'fr2.Open("select * from oc_order_total where order_id=" & Str(Adodc1.Recordset.Fields("order_id").Value) & "  and code='shipping'")
        '  238:        Dim R5 As New ADODB.Recordset
        If fr3data.Rows.Count > 0 Then
            poso = 1
            timm = fr3data.Rows(0).Item("Value").ToString / 1.24
            If timm > 0 Then
                kod = "ANTIK"
                onoma = "ANTIKATABOÀ«"
                SQL = ""
                SQL = "insert into PEGGTIM (APOT,ONOMA,HME,EIDOS,ATIM,PELKOD,KODE,POSO,TIMM,FPA,ID_NUM) VALUES ('1','" + onoma + "','" + SHMERA + "','e','"
                SQL = SQL & "a" & Mid(PELKOD, 2, 6) & "','"
                SQL = SQL & KPE & "','"
                SQL = SQL + kod + "',"
                SQL = SQL & Replace(VB6.Format(poso, "####.00"), ",", ".") & ","
                SQL = SQL & Replace(VB6.Format(timm, "####.00"), ",", ".") & ","
                SQL = SQL & "6,"
                SQL = SQL & Str(mID_NUM) & ")"
                AJ6 = AJ6 + poso * timm
                Gdb.Execute(SQL, nk)
                PrintLine(2, SQL)
            End If

            'End If
        End If

        '===================  ÙÔÍÔÈ ===================================================================================

        ExecuteSQLQuery("select * from oc_order_total where order_id=" & Str(SQLDT2.Rows(mRow).Item("order_id").ToString) & "  and code='tokoi'", fr3data)

        'fr2.Open("select * from oc_order_total where order_id=" & Str(Adodc1.Recordset.Fields("order_id").Value) & "  and code='shipping'")
        '  238:        Dim R5 As New ADODB.Recordset
        If fr3data.Rows.Count > 0 Then
            poso = 1
            timm = fr3data.Rows(0).Item("Value").ToString / 1.24
            If timm > 0 Then
                kod = "TOKOI"
                onoma = "‘œ œ…"
                SQL = ""
                SQL = "insert into PEGGTIM (APOT,ONOMA,HME,EIDOS,ATIM,PELKOD,KODE,POSO,TIMM,FPA,ID_NUM) VALUES ('1','" + onoma + "','" + SHMERA + "','e','"
                SQL = SQL & "a" & Mid(PELKOD, 2, 6) & "','"
                SQL = SQL & KPE & "','"
                SQL = SQL + kod + "',"
                SQL = SQL & Replace(VB6.Format(poso, "####.00"), ",", ".") & ","
                SQL = SQL & Replace(VB6.Format(timm, "####.00"), ",", ".") & ","
                SQL = SQL & "6,"
                SQL = SQL & Str(mID_NUM) & ")"
                AJ6 = AJ6 + poso * timm
                Gdb.Execute(SQL, nk)
                PrintLine(2, SQL)
            End If

            'End If
        End If









        '268:        fr2.Close()



        '270 fr2.open "select * from oc_order_total where orders_id=" + Str(Adodc1.Recordset("order_id")) + " and code='shipping'", GdbS, adOpenDynamic, adLockOptimistic
        '
        '272    If Not fr2.EOF Then
        '274       poso = 1
        ''276       timm = fr2!Value / 1.23
        ''278       kod = "METAF"
        ''280       onoma = "≈Œœƒ¡ Ã≈‘¡÷œ—¡”"
        ''282       SQL = ""
        ''284       SQL = "insert into PEGGTIM (APOT,ONOMA,HME,EIDOS,ATIM,PELKOD,KODE,POSO,TIMM,FPA,ID_NUM) VALUES ('1','" + onoma + "','" + SHMERA + "','e','"
        ''286       SQL = SQL + "a" + Mid(PELKOD, 2 ,6) + "','"
        ''288       SQL = SQL + KPE + "','"
        ''290       SQL = SQL + kod + "',"
        ''292       SQL = SQL + Replace(Format(poso, "####.00"), ",", ".") + ","
        ''294       SQL = SQL + Replace(Format(timm, "####.00"), ",", ".") + ","
        ''296       SQL = SQL + "2 ,"
        ''          SQL = SQL + Str(mID_NUM) + ")"
        ''298       aj2 = aj2 + poso * timm
        ''300       Gdb.Execute SQL, nk
        '       End If
        '302 fr2.Close

304:    'UPGRADE_WARNING: Couldn't resolve default property of object AJ7. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object AJ6. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object aj1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object aji. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        aji = aj1 * 1.13 + AJ6 * 1.24 + AJ7 * 1.17 ' + aj4 * 1.065
306:    'UPGRADE_WARNING: Couldn't resolve default property of object AJ3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AJ3 = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object AJ5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AJ5 = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object aji. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If aji = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object PELKOD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileClose(2)
            MsgBox("ƒ≈Õ ≈ ‘≈À≈”‘« ≈ ")
            Gdb.Execute("DELETE from PEGGTIM where  ATIM='a" & Mid(PELKOD, 2, 6) & "' AND HME='" & SHMERA & "'")


        End If


308:    SQL = ""
310:    SQL = "insert into PTIM (HME,EIDOS,ATIM,KPE,AJ1,AJ2,AJ3,AJ4,AJ5,AJ6,AJ7,AJI) VALUES ('" & SHMERA & "','e','"
312:    'UPGRADE_WARNING: Couldn't resolve default property of object PELKOD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SQL = SQL & "a" & Mid(PELKOD, 2, 6) & "','"
314:    SQL = SQL & KPE & "',"
316:    'UPGRADE_WARNING: Couldn't resolve default property of object aj1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SQL = SQL & Replace(VB6.Format(aj1, "####.00"), ",", ".") & ","
318:    'UPGRADE_WARNING: Couldn't resolve default property of object aj2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SQL = SQL & Replace(VB6.Format(aj2, "####.00"), ",", ".") & ","
320:    'UPGRADE_WARNING: Couldn't resolve default property of object AJ3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SQL = SQL & Replace(VB6.Format(AJ3, "####.00"), ",", ".") & ","
322:    'UPGRADE_WARNING: Couldn't resolve default property of object aj4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SQL = SQL & Replace(VB6.Format(aj4, "####.00"), ",", ".") & ","
324:    'UPGRADE_WARNING: Couldn't resolve default property of object AJ5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SQL = SQL & Replace(VB6.Format(AJ5, "####.00"), ",", ".") & ","
        'UPGRADE_WARNING: Couldn't resolve default property of object AJ6. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SQL = SQL & Replace(VB6.Format(AJ6, "####.00"), ",", ".") & ","

        'UPGRADE_WARNING: Couldn't resolve default property of object AJ7. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SQL = SQL & Replace(VB6.Format(AJ7, "####.00"), ",", ".") & ","

326:    'UPGRADE_WARNING: Couldn't resolve default property of object aji. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SQL = SQL & Replace(VB6.Format(aji, "####.00"), ",", ".") & ")"
328:    Gdb.Execute(SQL, nk)


        'UPGRADE_WARNING: Couldn't resolve default property of object nk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nk = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object PELKOD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Gdb.Execute("DELETE from PEGGTIM where  ATIM='a" & Mid(PELKOD, 2, 6) & "' AND HME='" & SHMERA & "'")
            'UPGRADE_WARNING: Couldn't resolve default property of object PELKOD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Gdb.Execute("DELETE from PTIM where  ATIM='a" & Mid(PELKOD, 2, 6) & "' AND HME='" & SHMERA & "'")
            FileClose(2)
            MsgBox("ƒ≈Õ ≈ ‘≈À≈”‘« ≈ ")

        End If






        RT.Open("SELECT MAX(ID_NUM) FROM PTIM", Gdb, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Dim m2ID_NUM As Integer

        m2ID_NUM = RT.Fields(0).Value
        Gdb.Execute("UPDATE PEGGTIM SET ID_NUM =" & Str(m2ID_NUM) & " WHERE ID_NUM=" & Str(mID_NUM))

        Me.Text = m_kodeid & " ÂÍÙÂÎ›ÛÙÁÍÂ " & VB6.Format(Now, "HH:MM")

        'MsgBox "ok"
        FileClose(2)

330:    'EXEC.Enabled = False

        '<EhFooter>
        Exit Sub

EXEC_Click_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.EXEC_Click " & "at line " & Erl())
        Resume Next
        '</EhFooter>
    End Sub

    Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        Dim c As String
        c = Text1.Text

        If Len(c) < 2 Then
            Exit Sub
        End If

        c = VB.Left(c, Len(c) - 1)
        MT = New DataTable

        ExecuteSQLQuery("SELECT  * FROM oc_order where order_id in(" & c & ") ", SQLDT2)

        ' Adodc1.RecordSource = "SELECT  * FROM oc_order where order_id in(" & c & ") "
        ' Adodc1.Refresh()
        Dim K As Integer


100:    For GGK As Integer = 0 To SQLDT2.Rows.Count - 1
102:        ' EXEC_Click(EXEC, New System.EventArgs())
104:        ' Adodc1.Recordset.MoveNext()
            update_order(GGK)
        Next

106:    MsgBox("œ ")



















        '        '<EhHeader>
        '        On Error GoTo Command3_Click_Err
        '        '</EhHeader>
        '100         EYRESH2_EKKREMON
        '        '<EhFooter>
        '        Exit Sub
        '
        'Command3_Click_Err:
        '      SAVE_ERROR Err.Description & vbCrLf & _
        ''               "in YOUHOU.Form1.Command3_Click " & _
        ''               "at line " & Erl
        '        Resume Next
        '</EhFooter>
    End Sub

    'UPGRADE_WARNING: Event f100.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub f100_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles f100.CheckStateChanged
        '<EhHeader>
        On Error GoTo f100_Click_Err
        '</EhHeader>
100:    If f100.CheckState = System.Windows.Forms.CheckState.Checked Then
102:        mF100 = 1000
        Else
104:        mF100 = 50
        End If

        '<EhFooter>
        Exit Sub

f100_Click_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.f100_Click " & "at line " & Erl())
        Resume Next
        '</EhFooter>
    End Sub

    Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'gdb local database
        'conn is web connection




        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        Dim SURL As Object

100:    mF100 = 100


102:    FileOpen(1, "C:\MERCVB\MERCPATH.TXT", OpenMode.Input)
104:    gDir = LineInput(1)
106:    gConnect = LineInput(1)
108:    FileClose(1)

        'host
        'db44.grserver.gr:3306
        'Database Name
        'web88_youdb
        'User Name
        'web88_youdbuser
        'Password
        'youhou!@#$



        '185.4.134.44

110:    '  gcon2 = "Provider=MSDASQL;DRIVER={MySQL ODBC 5.3 Unicode Driver};SERVER=185.4.134.44;PORT =3306;DATABASE=web88_youdb;UID=web88_youdbuser;Password=youhou!@#$"


112:    'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Len(Dir("C:\MERCVB\SITE.TXT")) > 0 Then
114:        FileOpen(1, "C:\MERCVB\SITE.TXT", OpenMode.Input)
            '    Line Input #1, gcon2
116:        FileClose(1)
        Else
118:        FileOpen(1, "C:\MERCVB\SITE.TXT", OpenMode.Output)
120:        PrintLine(1, gcon2)
122:        FileClose(1)

        End If

        'gcon2 = "DSN=youh2"

        connect_MySQL()

124:    Gdb.Open(Trim(gConnect)) '"DSN=MPOYG;UID=sa;pwd=gboug21275" '

        '‰ÔıÎÂıÂÈ ·ÎÎ· ‚„·ÊÂÈ Í·Ò·„ÍÈÔÊ·ÍÈ· ÛÙ· ÂÎÎÁÌÈÍ· „Ò·ÏÏ·Ù·
        'gcon2 = "DRIVER={MySQL odbc 3.51 Driver};SERVER=185.4.134.44;PORT =3306;DATABASE=web88_youdb;UID=web88_youdbuser;Password=youhou!@#$"


        ' gcon2 = "DSN=SITE64;UID=web88_youdbuser;pwd=youhou!@#$;"

126:    ' GdbS.Open(gcon2)
        'Adodc1.ConnectionString = gcon2
        ' Adodc1.RecordSource = "SELECT  * FROM oc_order where order_status_id=1 ORDER BY date_added DESC limit 10"
        ' Adodc1.Refresh()
        '  DataGrid1.DataSource = Adodc1



        ' SURL = "http://www.toys-shop.gr/admin"
        ' SURL = "http://www.toys-shop.gr/admin/orders.php"

        ' DataGrid1.DataSource = "Adodc1"
128:    'Adodc1.ConnectionString = gcon2



        ' date_added AS HMEPOMHNIA
        'order_id,date_added,firstname,lastname
130:    'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fSQL = "SELECT  * FROM oc_order where order_status_id=1 ORDER BY date_added DESC limit " & Str(mF100)
132:    'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '   Adodc1.RecordSource = fSQL
134:    '  Adodc1.Refresh()


        'SELECT `oc_order_status`.`order_status_id`,
        '    `oc_order_status`.`language_id`,
        '    `oc_order_status`.`name`
        'FROM `web88_youdb`.`oc_order_status`;
        Dim mfr As New DataTable
136:    ExecuteSQLQuery("select * from oc_order_status WHERE language_id=2 order by order_status_id ", mfr)
        Dim k As Integer
138:    For k = 0 To mfr.Rows.Count - 1
140:        Combo1.Items.Add(VB6.Format(mfr.Rows(k).Item("order_status_id").ToString, "##") & " ." + mfr.Rows(k).Item("Name").ToString)
142:
        Next
144:    Combo1.Text = VB6.GetItemString(Combo1, 0)



146:    'fr.Close()

        ' Exit Sub


        'where products_model='9316188036856'
148:    ExecuteSQLQuery("SELECT * FROM oc_product limit 1 ", mfr)
        '
        '
        '' deixnei ola ta pedia enos pinaka
150:    FileOpen(1, "C:\MERCVB\products.TXT", OpenMode.Output)
152:    For k = 0 To mfr.Columns.Count - 1
154:        PrintLine(1, mfr.Columns(k).ColumnName.ToString)
            '  Print #1, fr(k)
        Next

156:    FileClose(1)


        ExecuteSQLQuery("SELECT * FROM oc_order limit 1 ", mfr)
        '
        '
        '' deixnei ola ta pedia enos pinaka
        : FileOpen(1, "C:\MERCVB\orders.TXT", OpenMode.Output)
        For k = 0 To mfr.Columns.Count - 1
            PrintLine(1, mfr.Columns(k).ColumnName.ToString)
            '  Print #1, fr(k)
        Next

        FileClose(1)





        'ot_fixed_payment_chg
        'ot_shipping
        Exit Sub


        Dim r As New ADODB.Recordset
158:    r.Open("select  count(*) FROM orders where orders_status=" & VB.Left(Combo1.Text, 2) & " ORDER BY date_purchased DESC limit  " & Str(mF100), GdbS, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
160:    Me.Text = "–¡—¡√√≈À…≈” " & Str(r.Fields(0).Value)
162:    r.Close()

        'r.open "select  * FROM orders where orders_status=" + Left(Combo1.Text, 2) + " ORDER BY date_purchased DESC limit  " + Str(mF100), GdbS, adOpenDynamic, adLockOptimistic
        'r.Close






        '<EhFooter>
        Exit Sub

Form_Load_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.Form_Load " & "at line " & Erl())
        Resume Next
        '</EhFooter>
    End Sub
    Function FETES_DELIM(ByRef LINE As Object, ByRef elem As Object) As Object
        '<EhHeader>
        On Error GoTo FETES_DELIM_Err
        '</EhHeader>
        Dim KL, KE As Object

        On Error GoTo MHNYMA
        '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
100:    For KE = 1 To 30
102:        'UPGRADE_WARNING: Couldn't resolve default property of object elem(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            elem(KE) = ""
        Next

104:    'UPGRADE_WARNING: Couldn't resolve default property of object KL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KL = 0 ' metraei xaraktires
106:    'UPGRADE_WARNING: Couldn't resolve default property of object KE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        KE = 0 ' metritis toy pinaka  ELEMENT
        Do
108:        'UPGRADE_WARNING: Couldn't resolve default property of object KE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            KE = KE + 1 ' metritis toy pinaka  ELEMENT
110:        'UPGRADE_WARNING: Couldn't resolve default property of object KL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            KL = KL + 1 ' metraei xaraktires

112:        'UPGRADE_WARNING: Couldn't resolve default property of object KL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object LINE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Do While Mid(LINE, KL, 1) <> Chr(13) ' tab
114:            'UPGRADE_WARNING: Couldn't resolve default property of object KL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object LINE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object elem(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                elem(KE) = elem(KE) + Mid(LINE, KL, 1)
116:            'UPGRADE_WARNING: Couldn't resolve default property of object KL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                KL = KL + 1 ' metraei xaraktires
118:            'UPGRADE_WARNING: Couldn't resolve default property of object KL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If KL > Len(LINE) Then Exit Do
            Loop
            '  KL = KL + 1 ' „È· Ì· ÂÒ·ÛÂÈ ÙÔ chr(10)

120:        'UPGRADE_WARNING: Couldn't resolve default property of object KL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If KL > Len(LINE) Then Exit Do

122:        'UPGRADE_WARNING: Couldn't resolve default property of object KL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Loop Until KL >= Len(LINE) 'OLO TO MHKOS THS GRAMMHS

124:    'UPGRADE_WARNING: Couldn't resolve default property of object FETES_DELIM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FETES_DELIM = 0
        Exit Function

MHNYMA:
        'HandleError "Par1:Fetesdelim"
126:    Resume Next

        '<EhFooter>
        Exit Function

FETES_DELIM_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.FETES_DELIM " & "at line " & Erl())
        Resume Next
        '</EhFooter>
    End Function

    Private Sub List1_DblClick()
        Dim List1 As Object
        '<EhHeader>
        On Error GoTo List1_DblClick_Err
        '</EhHeader>
100:    'UPGRADE_WARNING: Couldn't resolve default property of object List1.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        List1.Left = 500
102:    'UPGRADE_WARNING: Couldn't resolve default property of object List1.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        List1.Width = 915
        '<EhFooter>
        Exit Sub

List1_DblClick_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.List1_DblClick " & "at line " & Erl())
        Resume Next
        '</EhFooter>
    End Sub

    '=================================ORDERS ====================================
    'orders_id
    '23273
    'customers_id
    '10727
    'customers_name
    '√—«√œ—«” Ã¡—…¡Õœ”
    'customers_company
    '
    'customers_street_address
    '¡Õ‘…–¡—œ”
    'customers_suburb
    ' ¡Àœ’ƒ…¡
    'customers_city
    '¡Õ‘…–¡—œ”
    'customers_postcode
    '84007
    'customers_state
    ' ’ À¡ƒ≈”
    'customers_country
    'Greece
    'customers_telephone
    '6977895836
    'customers_email_address
    'gmarianos67@ yahoo.gr
    'customers_address_format_id
    '1
    'delivery_name
    '√—«√œ—«” Ã¡—…¡Õœ”
    'delivery_company
    '
    'delivery_street_address
    '¡Õ‘…–¡—œ”
    'delivery_suburb
    ' ¡Àœ’ƒ…¡
    'delivery_city
    '¡Õ‘…–¡—œ”
    'delivery_postcode
    '84007
    'delivery_state
    ' ’ À¡ƒ≈”
    'delivery_country
    'Greece
    'delivery_address_format_id
    '1
    'billing_name
    '√—«√œ—«” Ã¡—…¡Õœ”
    'billing_company
    '
    'billing_street_address
    '¡Õ‘…–¡—œ”
    'billing_suburb
    ' ¡Àœ’ƒ…¡
    'billing_city
    '¡Õ‘…–¡—œ”
    'billing_postcode
    '84007
    'billing_state
    ' ’ À¡ƒ≈”
    'billing_country
    'Greece
    'billing_address_format_id
    '1
    'payment_method
    '–ÎÁÒ˘Ïﬁ ÏÂ ÈÛÙ˘ÙÈÍﬁ Í‹ÒÙ· (VISA, MASTERCARD, DINERS CLUB) ÛÂ ·Ûˆ·Î›Ú ÂÒÈ‚‹ÎÎÔÌ SSL+Points
    'cc_type
    '
    'cc_owner
    '
    'cc_number
    '
    'cc_expires
    '
    'last_modified
    '1/10/2013 8:20:41 Ï
    'date_purchased
    '1/10/2013 7:35:58 Ï
    'orders_status
    '2
    'orders_date_finished
    'Null
    'currency
    'EUR
    'currency_value
    '1
    'invoice_company
    '
    'invoice_job
    '
    'invoice_address
    '
    'invoice_afm
    '
    'invoice_doy
    '
    'survey






    '"orders_products_id";"orders_id";"products_id";"products_model";"products_name";"products_price";"final_price";"products_tax";"products_quantity"
    '40;14;742;"5204275010251";"";"33.61";"33.61";"19.00";1
    '37;12;88;"400878904270";"";"84.03";"84.03";"19.00";1


    '"orders_status_id";"language_id";"orders_status_name";"public_flag";"downloads_flag"
    '1;1;"ÂÍÍÒÂÏﬁÚ";1;0
    '4;4;"·ÍıÒ˛ËÁÍÂ";1;0
    '2;1;"send";1;0
    '3;1;"·¸‰ÂÈÓÁ";1;0
    '1;4;"ÂÍÍÒÂÏﬁÚ";1;0
    '2;4;"·ÂÛÙ‹ÎÁ";1;0
    '3;4;"·¸‰ÂÈÓÁ";1;0
    '4;1;"cancelled";1;0
    '5;4;"ÂÒÈÛıÎÎÔ„ﬁ";1;0
    '5;1;"ÂÒÈÛıÎÎÔ„ﬁ";1;0
    '6;4;"Œ‹ÌËÁ";1;0
    '6;1;"Xanthi";1;0
    '7;4;"·Ì·Ï›ÌÂÙ·È";1;0
    '7;1;"upcoming";1;0
    '8;4;"ÒÔÚ ·Ò·„„ÂÎﬂ·";1;0
    '8;1;"to be ordered";1;0
    '9;4;"Œ‹ÌËÁ ÂÒÈÛıÎÎÔ„ﬁ";1;0
    '9;1;"Xanthi picking";1;0
    '10;4;"·Ì·Ï›ÌÂÙ·È Perego";1;0
    '10;1;"expected Perego";1;0
    '11;4;"·Ì·Ï›ÌÂÙ·È BebeStars";1;0
    '11;1;"expected BebeStars";1;0
    '12;4;"·Ì·Ï›ÌÂÙ·È Quinny";1;0
    '12;1;"expected Quinny";1;0
    '13;4;"Preparing [PayPal Standard]";0;0
    '13;1;"Preparing [PayPal Standard]";0;0
    '14;4;"·Ì·Ï›ÌÂÙ·È CHICCO";1;0
    '14;1;"";1;0
    ''
    Private Sub Text_Change()

    End Sub

    Private Sub Text_LostFocus()

        'fSQL = "SELECT  * FROM orders_total where orders_id=" + Text.Text
        'Adodc1.RecordSource = fSQL
        'Adodc1.Refresh





    End Sub

    Private Sub id2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles id2.Leave



        '</EhHeader>
        If Val(id2.Text) > 0 Then

100:        'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            fSQL = "SELECT  *  FROM oc_order where order_id>=" & Str(CDbl(id1.Text)) & " and order_id<=" & Str(CDbl(id2.Text)) & " and order_status_id=" & VB.Left(Combo1.Text, 2) & " ORDER BY date_added DESC limit  " & Str(mF100)
102:        'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'Adodc1.RecordSource = fSQL
104:        ' Adodc1.Refresh()

        End If







    End Sub

    'UPGRADE_WARNING: Event oles.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub oles_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles oles.CheckStateChanged
        '<EhHeader>
        On Error GoTo oles_Click_Err
        '</EhHeader>


100:    'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fSQL = "SELECT  *  FROM oc_order ORDER BY order_id  DESC limit  " & Str(mF100)
102:    'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Adodc1.RecordSource = fSQL
104:    ' Adodc1.Refresh()

        Dim r As New ADODB.Recordset
106:    r.Open("SELECT  COUNT(*) FROM oc_order ORDER BY order_id  DESC limit  " & Str(mF100), GdbS, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
108:    Me.Text = "–¡—¡√√≈À…≈” " & Str(r.Fields(0).Value)

110:    r.Close()

        '<EhFooter>
        Exit Sub

oles_Click_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.oles_Click " & "at line " & Erl())
        Resume Next
        '</EhFooter>
    End Sub

    Private Sub Text1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.Leave
        '<EhHeader>
        On Error GoTo Text1_LostFocus_Err
        '</EhHeader>
100:    'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fSQL = "SELECT  * FROM oc_order where order_id=" & Text1.Text & "   limit  " & Str(mF100)
102:    'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Adodc1.RecordSource = fSQL
104:    'Adodc1.Refresh()
        '<EhFooter>
        Exit Sub

Text1_LostFocus_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.Text1_LostFocus " & "at line " & Erl())
        Resume Next
        '</EhFooter>
    End Sub







    Sub SAVE_ERROR(ByRef COMMENT As Object)
        '<EhHeader>
        On Error GoTo SAVE_ERROR_Err
        '</EhHeader>

        On Error Resume Next

        'SAVE_ERROR Err.Description & " in Project1.Form1.cmdCommand2_Click " & " at line " & Erl
        Dim f As Short

100:    f = FreeFile()
102:    FileOpen(f, "C:\MERCVB\ERR.TXT", OpenMode.Append)
104:    'UPGRADE_WARNING: Couldn't resolve default property of object COMMENT. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        WriteLine(f, VB6.Format(Now, "DD/MM/YYYY HH:MM") + COMMENT)

106:    FileClose(f)

        '<EhFooter>
        Exit Sub

SAVE_ERROR_Err:
        SAVE_ERROR(Err.Description & vbCrLf & "in YOUHOU.Form1.SAVE_ERROR " & "at line " & Erl())
        Resume Next
        '</EhFooter>
    End Sub

    Private Sub DataGrid1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGrid1.CellClick
        If DataGrid1.SelectedRows.Count = 0 Then
            ' Exit Sub

        End If
        If InStr(Text1.Text, DataGrid1.CurrentRow.Cells(0).Value.ToString) > 0 Then
            Me.Text = "ı·Ò˜ÂÈ Á‰Á"
        Else
            Text1.Text = Text1.Text & DataGrid1.CurrentRow.Cells(0).Value.ToString & ","
            Me.Text = ""
        End If
    End Sub

    Private Sub DataGrid1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGrid1.CellContentClick
        If DataGrid1.SelectedRows.Count = 0 Then
            Exit Sub

        End If
        If InStr(Text1.Text, DataGrid1.SelectedRows(0).Cells(0).Value.ToString) > 0 Then
            Me.Text = "ı·Ò˜ÂÈ Á‰Á"
        Else
            Text1.Text = Text1.Text & DataGrid1.SelectedRows(0).Cells(0).Value.ToString & ","
            Me.Text = ""
        End If
    End Sub

    Private Sub DataGrid1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGrid1.RowHeaderMouseClick

    End Sub

    Private Sub –·Ò·„„ÂÎÈÂÚ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles –·Ò·„„ÂÎÈÂÚ.Click
       

100:    'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fSQL = "SELECT  *  FROM oc_order where date_added>='" + Format(D1.Value, "yyyy/MM/dd") + "' and date_added<='" + Format(D2.Value, "yyyy/MM/dd") + "'  ORDER BY date_added DESC limit 500 " '& Str(mF100)
        'fSQL = "SELECT  *  FROM oc_order where order_status_id=" & VB.Left(Combo1.Text, 2) & " ORDER BY date_added DESC limit  " & Str(mF100)
102:    'UPGRADE_WARNING: Couldn't resolve default property of object fSQL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Adodc1.RecordSource = fSQL
104:    ' Adodc1.Refresh()

        'Exit Sub


        data = New DataTable

        da = New MySqlDataAdapter(fSQL, conn)
        ' cb = New MySqlCommandBuilder(da)

        da.Fill(data)

        DataGrid1.DataSource = data

        ' DataGrid1.ReBind()
        DataGrid1.Refresh()

        DataGrid1.DataSource = data

        ExecuteSQLQuery(fSQL, SQLDT2)

        ' Adodc1.RecordSource = "SELECT  * FROM oc_order where order_id in(" & c & ") "
        ' Adodc1.Refresh()
        Dim K As Integer


        For GGK As Integer = 0 To SQLDT2.Rows.Count - 1
            ' EXEC_Click(EXEC, New System.EventArgs())
            ' Adodc1.Recordset.MoveNext()
            update_order(GGK)
        Next


        '        Dim r As New ADODB.Recordset
        '106:    r.Open("select  count(*) FROM oc_order where order_status_id=" & VB.Left(Combo1.Text, 2), GdbS, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        '108:    Me.Text = "–¡—¡√√≈À…≈” " & Str(r.Fields(0).Value)
        '110:    r.Close()



    End Sub


    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        Dim c As String
        c = Text1.Text

        If Len(c) < 2 Then
            Exit Sub
        End If

        c = VB.Left(c, Len(c) - 1)








        '--------------------------------------------------------------------------------------
        Dim fsql As String = "SELECT  * FROM oc_order where order_id in(" & c & ") "
        data = New DataTable

        da = New MySqlDataAdapter(fsql, conn)
        ' cb = New MySqlCommandBuilder(da)

        da.Fill(data)

        DataGrid1.DataSource = data

        ' DataGrid1.ReBind()
        DataGrid1.Refresh()

        DataGrid1.DataSource = data
        '--------------------------------------------------------------------------------------




        MT = New DataTable

        ExecuteSQLQuery(fsql, SQLDT2)

        ' Adodc1.RecordSource = "SELECT  * FROM oc_order where order_id in(" & c & ") "
        ' Adodc1.Refresh()
        Dim K As Integer


100:    For GGK As Integer = 0 To SQLDT2.Rows.Count - 1
102:        ' EXEC_Click(EXEC, New System.EventArgs())
104:        ' Adodc1.Recordset.MoveNext()
            update_order(GGK)
        Next

106:    MsgBox("œ ")









    End Sub



End Class