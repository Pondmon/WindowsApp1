﻿Imports System.Data.SqlClient
Imports System.Drawing.Imaging
Imports IDAutomation.Windows.Forms.LinearBarCode
Imports System.Drawing.Printing
Imports System.Configuration
Imports GenCode128
Imports MySql.Data.MySqlClient

Public Class Form2
    Private WithEvents pdPrint As PrintDocument
    Private PrintDocType As String = "Barcode"
    Private StrPrinterName As String = "Canon MP280 series"
    Public old_barcode_id As String

    Dim pbImage2 As New PictureBox
    Private bmp As Bitmap
    Dim document_type As String = "คำร้องขอลงทะเบียนเพิ่ม/เปลี่ยนกลุ่ม กรณีกลุ่มเต็ม/ลดรายวิชา"
    Dim barcode_id As String

    Dim prtdoc As New PrintDocument
    Dim strDefaultPrinter As String = prtdoc.PrinterSettings.PrinterName

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Clear text box
        Text_ID.Text = ""
        Text_Name.Text = ""
        Text_SUB.Text = ""
        Text_IDSUB.Text = ""
        Text_School.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        Text_INS.Text = ""
        TextBox11.Text = ""
        TextBox14.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox4.Text = ""

    End Sub
    Public Sub Print_Document()

        Hiden()

        bmp = New Bitmap(11906, 16838)
        Dim G As Graphics = Graphics.FromImage(bmp)

        Panel1.DrawToBitmap(bmp, Panel1.ClientRectangle)
        G.Dispose()

        PrintDocument1.DefaultPageSettings.PaperSize = New PaperSize("210 x 297 mm", 815, 1200)
        PrintPreviewDialog1.Document = PrintDocument1

        PrintPreviewDialog1.ShowDialog()
        'PrintDocument1.Print()
    End Sub

    Public Sub Nomal()
        Text_ID.BorderStyle = BorderStyle.Fixed3D
        Text_Name.BorderStyle = BorderStyle.Fixed3D
        Text_School.BorderStyle = BorderStyle.Fixed3D
        TextBox6.BorderStyle = BorderStyle.Fixed3D
        TextBox7.BorderStyle = BorderStyle.Fixed3D
        TextBox8.BorderStyle = BorderStyle.Fixed3D
        Text_INS.BorderStyle = BorderStyle.Fixed3D
        TextBox11.BorderStyle = BorderStyle.Fixed3D
        TextBox14.BorderStyle = BorderStyle.Fixed3D
    End Sub

    Public Sub Hiden()
        Text_ID.BorderStyle = BorderStyle.None
        Text_Name.BorderStyle = BorderStyle.None
        Text_School.BorderStyle = BorderStyle.None
        TextBox6.BorderStyle = BorderStyle.None
        TextBox7.BorderStyle = BorderStyle.None
        TextBox8.BorderStyle = BorderStyle.None
        Text_INS.BorderStyle = BorderStyle.None
        TextBox11.BorderStyle = BorderStyle.None
        TextBox14.BorderStyle = BorderStyle.None

    End Sub

    Public Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        e.Graphics.DrawImage(bmp, 0, 0)
    End Sub

    Public Sub GEN_Barcode()
        ' Gen Barcode
        'ID Automation
        'Free only with the Code39 and Code39Ext
        Dim NewBarcode As IDAutomation.Windows.Forms.LinearBarCode.Barcode = New Barcode()

        Dim MyDate As String = DateTime.Now.ToString("HHmmddMMyy")
        'date_time_now = DateTime.Now.ToString
        barcode_id = Text_ID.Text.ToString() & MyDate
        'MsgBox(barcode_id)

        NewBarcode.DataToEncode = barcode_id  'Input of textbox to generate barcode 

        NewBarcode.SymbologyID = Symbologies.Code39
        NewBarcode.Code128Set = Code128CharacterSets.A
        NewBarcode.RotationAngle = RotationAngles.Zero_Degrees
        NewBarcode.RefreshImage()
        NewBarcode.Resolution = Resolutions.Screen
        NewBarcode.ResolutionCustomDPI = 96
        NewBarcode.RefreshImage()

        NewBarcode.SaveImageAs("SavedBarcode.Jpeg", System.Drawing.Imaging.ImageFormat.Jpeg)
        NewBarcode.Resolution = Resolutions.Printer

        PictureBox1.Image = Image.FromFile(Application.StartupPath & "\" & "SavedBarcode.Jpeg")

        'Barcode using the GenCode128
        'Dim myimg As Image = Code128Rendering.MakeBarcodeImage(Text_ID.Text.ToString(), 1, False)
        'PictureBox2.Image = myimg
        'pbImage2.Image = myimg
        'Barcode using the GenCode128
    End Sub

    Public Sub save_SQL()
        ' Dim Query As String = "INSERT INTO `inform_std` (STD_ID,STD_NAME,STD_SUBJECT,STD_IDSub,STD_ADVISOR,STD_LECTURER,DATE,STD_STATUS,STD_TYPE) VALUES ('B12344','ปอร์ด','ไมโครเวฟ','123456','John','Jame','2019-4-22','...','12')"
        ' Dim  As String = "SET character_set_connection=utf8"
        'TextBox1.Text = TextBox1.Text & "Qury Text:" & Query & vbCrLf
        ' Query = " (eid,name,surname,age) values ('" & TextBox_Eid.Text & "','" & TextBox_Name.Text & "','" & TextBox_SName.Text & "','" & TextBox_Age.Text & "')"

        Dim conn As New MySql.Data.MySqlClient.MySqlConnection
        Dim myConnectionString As String

        myConnectionString = "server='178.128.63.128';" _
                       & "uid = TCEformSQL;" _
                      & "pwd='TCEsut1234*';" _
                      & "database=SUT_Student_Project;" _
                     & "charset=utf8;"

        'myConnectionString = "server='127.0.0.1';" _
        '              & "uid = root;" _
        '             & "pwd='';" _
        '             & "database=student_database;" _
        '            & "charset=utf8;"

        Try
            conn.ConnectionString = myConnectionString
            conn.Open()
            'MsgBox("SQL Database Connect OK!")

            Dim Query_CMD As String
            'Query_CMD = "INSERT INTO `TCE_database`(`std_id`, `std_name`, `From_type`) VALUES ('B123456','ศิวพร ใหญ่กลาง','02')"
            Query_CMD = "INSERT INTO `inform_std`( `STD_ID`, `STD_NAME`, `STD_SUBJECT`, `STD_IDSub`, `STD_LECTURER`,`STD_Faculty`, `STD_School`, `STD_STATUS`, `STD_TYPE`,`STD_Barcode`) VALUES ('" & Text_ID.Text & "','" & Text_Name.Text & "','" & Text_SUB.Text & "','" & Text_IDSUB.Text & "','" & ComboBox4.Text & "','" & Text_INS.Text & "','" & Text_School.Text & "','ส่งเอกสาร','" & document_type & "','" & barcode_id & "')"

            Dim COMMAND As MySqlCommand
            COMMAND = New MySqlCommand(Query_CMD, conn)

            Dim READER As MySqlDataReader
            READER = COMMAND.ExecuteReader

            MessageBox.Show("บันทึกข้อมูลเรียบร้อยแล้ว", "แจ้งเตือน")
            conn.Close()

        Catch ex As MySql.Data.MySqlClient.MySqlException
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try
    End Sub
    Public Sub Update_SQL()
        ' Dim Query As String = "INSERT INTO `inform_std` (STD_ID,STD_NAME,STD_SUBJECT,STD_IDSub,STD_ADVISOR,STD_LECTURER,DATE,STD_STATUS,STD_TYPE) VALUES ('B12344','ปอร์ด','ไมโครเวฟ','123456','John','Jame','2019-4-22','...','12')"
        ' Dim  As String = "SET character_set_connection=utf8"
        'TextBox1.Text = TextBox1.Text & "Qury Text:" & Query & vbCrLf
        ' Query = " (eid,name,surname,age) values ('" & TextBox_Eid.Text & "','" & TextBox_Name.Text & "','" & TextBox_SName.Text & "','" & TextBox_Age.Text & "')"
        Dim NewBarcode As IDAutomation.Windows.Forms.LinearBarCode.Barcode = New Barcode()

        Dim MyDate As String = DateTime.Now.ToString("HHmmddMMyy")
        'date_time_now = DateTime.Now.ToString
        barcode_id = Text_ID.Text.ToString() & MyDate
        'MsgBox(barcode_id) 

        NewBarcode.SymbologyID = Symbologies.Code39
        NewBarcode.Code128Set = Code128CharacterSets.A
        NewBarcode.RotationAngle = RotationAngles.Zero_Degrees
        NewBarcode.RefreshImage()
        NewBarcode.Resolution = Resolutions.Screen
        NewBarcode.ResolutionCustomDPI = 96
        NewBarcode.RefreshImage()

        NewBarcode.SaveImageAs("SavedBarcode.Jpeg", System.Drawing.Imaging.ImageFormat.Jpeg)
        NewBarcode.Resolution = Resolutions.Printer

        PictureBox1.Image = Image.FromFile(Application.StartupPath & "\" & "SavedBarcode.Jpeg")

        Dim conn As New MySql.Data.MySqlClient.MySqlConnection
        Dim myConnectionString As String

        myConnectionString = "server='178.128.63.128';" _
                       & "uid = TCEformSQL;" _
                      & "pwd='TCEsut1234*';" _
                      & "database=SUT_Student_Project;" _
                     & "charset=utf8;"

        'myConnectionString = "server='127.0.0.1';" _
        '              & "uid = root;" _
        '             & "pwd='';" _
        '             & "database=student_database;" _
        '            & "charset=utf8;"

        Try
            conn.ConnectionString = myConnectionString
            conn.Open()
            'MsgBox("SQL Database Connect OK!")

            Dim Query_CMD As String
            'UPDATE `inform_std` SET `STD_ID`=[value-1],`STD_NAME`=[value-2],`STD_SUBJECT`=[value-3],`STD_IDSub`=[value-4],`STD_ADVISOR`=[value-5],`STD_LECTURER`=[value-6],`STD_Professor`=[value-7],`STD_STATUS`=[value-8],`STD_TYPE`=[value-9],`STD_Barcode`=[value-10],`DATE`=[value-11] WHERE 1
            Query_CMD = "UPDATE `inform_std` SET `STD_ID`='" & Text_ID.Text & "',`STD_NAME`='" & Text_Name.Text & "',`STD_SUBJECT`='" & Text_SUB.Text & "',`STD_IDSub`='" & Text_IDSUB.Text & "',`STD_LECTURER`='" & ComboBox4.Text & "',`STD_Faculty`='" & Text_INS.Text & "', `STD_School`='" & Text_School.Text & "' WHERE `STD_Barcode`='" & old_barcode_id.ToString & "'"
            'MessageBox.Show(Query_CMD)

            Dim COMMAND As MySqlCommand
            COMMAND = New MySqlCommand(Query_CMD, conn)

            Dim READER As MySqlDataReader
            READER = COMMAND.ExecuteReader

            MessageBox.Show("บันทึกข้อมูลเรียบแล้ว", "แจ้งเตือน")
            conn.Close()

        Catch ex As MySql.Data.MySqlClient.MySqlException
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try

    End Sub

    Private Sub btBack_Click(sender As Object, e As EventArgs) Handles btBack.Click
        Form1.Show()
        Me.Close()
    End Sub

    Private Sub btSave_Click(sender As Object, e As EventArgs) Handles btSave.Click

        'Label39.Text = date_time_now

        GEN_Barcode()

        Dim result As DialogResult
        result = MessageBox.Show("ต้องการบันทึกข้อมูลหรือไม่ ??", "แจ้งเตือน", MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Yes Then
            save_SQL()
            Print_Document()

            Text_ID.Text = ""
            Text_Name.Text = ""
            Text_SUB.Text = ""
            Text_IDSUB.Text = ""
            Text_School.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            Text_INS.Text = ""
            TextBox11.Text = ""
            TextBox14.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox4.Text = ""
            PictureBox1.Image = Nothing
            Nomal()

        ElseIf result = DialogResult.No Then
            Text_ID.Text = ""
            Text_Name.Text = ""
            Text_SUB.Text = ""
            Text_IDSUB.Text = ""
            Text_School.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            Text_INS.Text = ""
            TextBox11.Text = ""
            TextBox14.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox4.Text = ""
            PictureBox1.Image = Nothing

        ElseIf result = DialogResult.Cancel Then
            PictureBox1.Image = Nothing
        End If

    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click

        Insert_Edit.form_ID_callback = "Form_2"
        Insert_Edit.Show()

    End Sub

    Private Sub Button_Edit_Click(sender As Object, e As EventArgs) Handles Button_Edit.Click
        Update_SQL()
        Print_Document()
    End Sub

End Class



