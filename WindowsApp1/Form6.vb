Imports System.Drawing.Imaging
Imports IDAutomation.Windows.Forms.LinearBarCode
Imports System.Drawing.Printing
Imports System.Configuration
Imports GenCode128
Imports MySql.Data.MySqlClient

Public Class Form6
    Private WithEvents pdPrint As PrintDocument
    Private PrintDocType As String = "Barcode"
    Private StrPrinterName As String = "Canon MP280 series"
    Public old_barcode_id As String

    Dim pbImage2 As New PictureBox
    Private bmp As Bitmap
    Dim document_type As String = "คำร้องขอลาระหว่างเรียน"
    Dim barcode_id As String

    Dim prtdoc As New PrintDocument
    Dim strDefaultPrinter As String = prtdoc.PrinterSettings.PrinterName

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Clear text box
        Text_ID.Text = ""
        Text_Name.Text = ""
        Text_SUB.Text = ""
        Text_IDSUB.Text = ""
        Text_S.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox22.Text = ""
        TextBox23.Text = ""
        TextBox24.Text = ""
        TextBox25.Text = ""
        TextBox26.Text = ""
        TextBox27.Text = ""
        Text_School.Text = ""
        Text_INS.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""
        ComboBox6.Text = ""
        ComboBox7.Text = ""
        ComboBox8.Text = ""
        ComboBox9.Text = ""

    End Sub

    Public Sub Print_Document()

        Hiden()

        bmp = New Bitmap(11906, 16838)
        Dim G As Graphics = Graphics.FromImage(bmp)

        Panel1.DrawToBitmap(bmp, Panel1.ClientRectangle)
        G.Dispose()

        PrintDocument1.DefaultPageSettings.PaperSize = New PaperSize("210 x 297 mm", 800, 1125)
        PrintPreviewDialog1.Document = PrintDocument1

        PrintPreviewDialog1.ShowDialog()
        'PrintDocument1.Print()
    End Sub
    Public Sub Nomal()
        Text_ID.BorderStyle = BorderStyle.Fixed3D
        Text_Name.BorderStyle = BorderStyle.Fixed3D
        Text_S.BorderStyle = BorderStyle.Fixed3D
        TextBox4.BorderStyle = BorderStyle.Fixed3D
        TextBox5.BorderStyle = BorderStyle.Fixed3D
        TextBox12.BorderStyle = BorderStyle.Fixed3D
        TextBox13.BorderStyle = BorderStyle.Fixed3D
        TextBox14.BorderStyle = BorderStyle.Fixed3D
        TextBox15.BorderStyle = BorderStyle.Fixed3D
        TextBox16.BorderStyle = BorderStyle.Fixed3D
        TextBox17.BorderStyle = BorderStyle.Fixed3D
        TextBox18.BorderStyle = BorderStyle.Fixed3D
        TextBox22.BorderStyle = BorderStyle.Fixed3D
        TextBox23.BorderStyle = BorderStyle.Fixed3D
        TextBox24.BorderStyle = BorderStyle.Fixed3D
        TextBox25.BorderStyle = BorderStyle.Fixed3D
        TextBox26.BorderStyle = BorderStyle.Fixed3D
        TextBox27.BorderStyle = BorderStyle.Fixed3D
        Text_INS.BorderStyle = BorderStyle.Fixed3D

    End Sub
    Public Sub Hiden()
        Text_ID.BorderStyle = BorderStyle.None
        Text_Name.BorderStyle = BorderStyle.None
        Text_S.BorderStyle = BorderStyle.None
        TextBox4.BorderStyle = BorderStyle.None
        TextBox5.BorderStyle = BorderStyle.None
        TextBox12.BorderStyle = BorderStyle.None
        TextBox13.BorderStyle = BorderStyle.None
        TextBox14.BorderStyle = BorderStyle.None
        TextBox15.BorderStyle = BorderStyle.None
        TextBox16.BorderStyle = BorderStyle.None
        TextBox17.BorderStyle = BorderStyle.None
        TextBox18.BorderStyle = BorderStyle.None
        TextBox22.BorderStyle = BorderStyle.None
        TextBox23.BorderStyle = BorderStyle.None
        TextBox24.BorderStyle = BorderStyle.None
        TextBox25.BorderStyle = BorderStyle.None
        TextBox26.BorderStyle = BorderStyle.None
        TextBox27.BorderStyle = BorderStyle.None
        Text_School.BorderStyle = BorderStyle.None
        Text_INS.BorderStyle = BorderStyle.None

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

    End Sub

    Public Sub save_SQL()

        Dim conn As New MySql.Data.MySqlClient.MySqlConnection
        Dim myConnectionString As String

        myConnectionString = "server='178.128.63.128';" _
                       & "uid = TCEformSQL;" _
                      & "pwd='TCEsut1234*';" _
                      & "database=SUT_Student_Project;" _
                     & "charset=utf8;"

        Try
            conn.ConnectionString = myConnectionString
            conn.Open()
            'MsgBox("SQL Database Connect OK!")

            Dim Query_CMD As String
            'Query_CMD = "INSERT INTO `TCE_database`(`std_id`, `std_name`, `From_type`) VALUES ('B123456','ศิวพร ใหญ่กลาง','02')"
            Query_CMD = "INSERT INTO `inform_std`( `STD_ID`, `STD_NAME`, `STD_SUBJECT`, `STD_IDSub`, `STD_ADVISOR`, `STD_LECTURER`,`STD_Faculty`, `STD_School`, `STD_STATUS`, `STD_TYPE`,`STD_Barcode`) VALUES ('" & Text_ID.Text & "','" & Text_Name.Text & "','" & Text_SUB.Text & "','" & Text_IDSUB.Text & "','" & ComboBox3.Text & "','" & Text_S.Text & "','" & Text_INS.Text & "','" & Text_School.Text & "','ส่งเอกสาร','" & document_type & "','" & barcode_id & "')"

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

    Public Sub Update_SQL()

        Dim NewBarcode As IDAutomation.Windows.Forms.LinearBarCode.Barcode = New Barcode()

        NewBarcode.DataToEncode = old_barcode_id  'Input of textbox to generate barcode 

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

        Try
            conn.ConnectionString = myConnectionString
            conn.Open()
            'MsgBox("SQL Database Connect OK!")

            Dim Query_CMD As String
            'UPDATE `inform_std` SET `STD_ID`=[value-1],`STD_NAME`=[value-2],`STD_SUBJECT`=[value-3],`STD_IDSub`=[value-4],`STD_ADVISOR`=[value-5],`STD_LECTURER`=[value-6],`STD_Professor`=[value-7],`STD_STATUS`=[value-8],`STD_TYPE`=[value-9],`STD_Barcode`=[value-10],`DATE`=[value-11] WHERE 1
            Query_CMD = "UPDATE `inform_std` SET `STD_ID`='" & Text_ID.Text & "',`STD_NAME`='" & Text_Name.Text & "',`STD_SUBJECT`='" & Text_SUB.Text & "',`STD_IDSub`='" & Text_IDSUB.Text & "',`STD_ADVISOR`='" & ComboBox3.Text & "',`STD_LECTURER`='" & Text_S.Text & "',`STD_Faculty`='" & Text_INS.Text & "', `STD_School`='" & Text_School.Text & "' WHERE `STD_Barcode`='" & old_barcode_id.ToString & "'"
            'MessageBox.Show(Query_CMD)

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
            Text_S.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox12.Text = ""
            TextBox13.Text = ""
            TextBox14.Text = ""
            TextBox15.Text = ""
            TextBox16.Text = ""
            TextBox17.Text = ""
            TextBox18.Text = ""
            TextBox22.Text = ""
            TextBox23.Text = ""
            TextBox24.Text = ""
            TextBox25.Text = ""
            TextBox26.Text = ""
            TextBox27.Text = ""
            Text_School.Text = ""
            Text_INS.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox3.Text = ""
            ComboBox4.Text = ""
            ComboBox5.Text = ""
            ComboBox6.Text = ""
            ComboBox7.Text = ""
            ComboBox8.Text = ""
            ComboBox9.Text = ""

            PictureBox1.Image = Nothing
            Nomal()

        ElseIf result = DialogResult.No Then
            Text_ID.Text = ""
            Text_Name.Text = ""
            Text_SUB.Text = ""
            Text_IDSUB.Text = ""
            Text_S.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox12.Text = ""
            TextBox13.Text = ""
            TextBox14.Text = ""
            TextBox15.Text = ""
            TextBox16.Text = ""
            TextBox17.Text = ""
            TextBox18.Text = ""
            TextBox22.Text = ""
            TextBox23.Text = ""
            TextBox24.Text = ""
            TextBox25.Text = ""
            TextBox26.Text = ""
            TextBox27.Text = ""
            Text_School.Text = ""
            Text_INS.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox3.Text = ""
            ComboBox4.Text = ""
            ComboBox5.Text = ""
            ComboBox6.Text = ""
            ComboBox7.Text = ""
            ComboBox8.Text = ""
            ComboBox9.Text = ""

            PictureBox1.Image = Nothing

        ElseIf result = DialogResult.Cancel Then
            PictureBox1.Image = Nothing
        End If

    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        Insert_Edit.form_ID_callback = "Form_6"
        Insert_Edit.Show()

    End Sub

    Private Sub Button_Edit_Click(sender As Object, e As EventArgs) Handles Button_Edit.Click
        Update_SQL()
        Print_Document()
    End Sub
End Class