Imports System.Data.SqlClient
Imports System.Drawing.Imaging
Imports IDAutomation.Windows.Forms.LinearBarCode
Imports System.Drawing.Printing
Imports System.Configuration
Imports GenCode128
Imports MySql.Data.MySqlClient

Public Class Form4

    Private WithEvents pdPrint As PrintDocument
    Private PrintDocType As String = "Barcode"
    Private StrPrinterName As String = "Canon MP280 series"

    Dim pbImage2 As New PictureBox
    Private bmp As Bitmap

    Dim prtdoc As New PrintDocument
    Dim strDefaultPrinter As String = prtdoc.PrinterSettings.PrinterName

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Clear text box
        Text_ID.Text = ""
        Text_NAME.Text = ""
        Text_SUB.Text = ""
        Text_IDSUB.Text = ""
        Text_ADV.Text = ""
        Text_IDSUB3.Text = ""
        Text_IDSUB2.Text = ""
        Text_IDSUB4.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox22.Text = ""
        TextBox23.Text = ""
        TextBox25.Text = ""
        TextBox26.Text = ""
        TextBox27.Text = ""
        TextBox32.Text = ""
        TextBox30.Text = ""
        TextBox31.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        PictureBox3.Image = Nothing
    End Sub
    Public Sub Print_Document()
        bmp = New Bitmap(Panel1.Width, Panel1.Height)
        Dim G As Graphics = Graphics.FromImage(bmp)

        Panel1.DrawToBitmap(bmp, Panel1.ClientRectangle)
        G.Dispose()

        PrintDocument1.DefaultPageSettings.PaperSize = New PaperSize("210 x 297 mm", 790, 1125)
        PrintPreviewDialog1.Document = PrintDocument1

        PrintPreviewDialog1.ShowDialog()
        'PrintDocument1.Print()
    End Sub
    Public Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        e.Graphics.DrawImage(bmp, 0, 0)
    End Sub
    Public Sub GEN_Barcode()
        ' Gen Barcode
        'ID Automation
        'Free only with the Code39 and Code39Ext
        Dim NewBarcode As IDAutomation.Windows.Forms.LinearBarCode.Barcode = New Barcode()

        NewBarcode.DataToEncode = Text_ID.Text.ToString() 'Input of textbox to generate barcode 

        NewBarcode.SymbologyID = Symbologies.Code39
        NewBarcode.Code128Set = Code128CharacterSets.A
        NewBarcode.RotationAngle = RotationAngles.Zero_Degrees
        NewBarcode.RefreshImage()
        NewBarcode.Resolution = Resolutions.Screen
        NewBarcode.ResolutionCustomDPI = 96
        NewBarcode.RefreshImage()

        NewBarcode.SaveImageAs("SavedBarcode.Jpeg", System.Drawing.Imaging.ImageFormat.Jpeg)
        NewBarcode.Resolution = Resolutions.Printer

        PictureBox3.Image = Image.FromFile(Application.StartupPath & "\" & "SavedBarcode.Jpeg")

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

        'myConnectionString = "server='178.128.63.128';" _
        '               & "uid = TCEformSQL;" _
        '              & "pwd='TCEsut1234*';" _
        '              & "database=SUT_Student_Project;" _
        '             & "charset=utf8;"

        myConnectionString = "server='127.0.0.1';" _
                       & "uid = root;" _
                     & "pwd='';" _
                     & "database=student_database;" _
                    & "charset=utf8;"

        Try
            conn.ConnectionString = myConnectionString
            conn.Open()
            'MsgBox("SQL Database Connect OK!")

            Dim Query_CMD As String
            'Query_CMD = "INSERT INTO `TCE_database`(`std_id`, `std_name`, `From_type`) VALUES ('B123456','ศิวพร ใหญ่กลาง','02')"
            Query_CMD = "INSERT INTO `inform_std`(`STD_ID`, `STD_NAME`, `STD_SUBJECT`, `STD_IDSub`, `STD_ADVISOR`, `STD_LECTURER`, `STD_STATUS`, `STD_TYPE`) VALUES ('" & Text_ID.Text & "','" & Text_NAME.Text & "','" & Text_SUB.Text & "','" & Text_IDSUB.Text & "','-','" & Text_ADV.Text & "','ส่งเอกสาร','03')"

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
        GEN_Barcode()

        Dim result As DialogResult
        result = MessageBox.Show("ต้องการบันทึกข้อมูลหรือไม่ ??", "แจ้งเตือน", MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Yes Then
            save_SQL()
            Print_Document()
            Text_ID.Text = ""
            Text_NAME.Text = ""
            Text_SUB.Text = ""
            Text_IDSUB.Text = ""
            Text_ADV.Text = ""
            Text_IDSUB3.Text = ""
            Text_IDSUB2.Text = ""
            Text_IDSUB4.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""
            TextBox12.Text = ""
            TextBox13.Text = ""
            TextBox14.Text = ""
            TextBox15.Text = ""
            TextBox16.Text = ""
            TextBox17.Text = ""
            TextBox18.Text = ""
            TextBox22.Text = ""
            TextBox23.Text = ""
            TextBox25.Text = ""
            TextBox26.Text = ""
            TextBox27.Text = ""
            TextBox32.Text = ""
            TextBox30.Text = ""
            TextBox31.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            PictureBox3.Image = Nothing

        ElseIf result = DialogResult.No Then
            Text_ID.Text = ""
            Text_NAME.Text = ""
            Text_SUB.Text = ""
            Text_IDSUB.Text = ""
            Text_ADV.Text = ""
            Text_IDSUB3.Text = ""
            Text_IDSUB2.Text = ""
            Text_IDSUB4.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""
            TextBox12.Text = ""
            TextBox13.Text = ""
            TextBox14.Text = ""
            TextBox15.Text = ""
            TextBox16.Text = ""
            TextBox17.Text = ""
            TextBox18.Text = ""
            TextBox22.Text = ""
            TextBox23.Text = ""
            TextBox25.Text = ""
            TextBox26.Text = ""
            TextBox27.Text = ""
            TextBox32.Text = ""
            TextBox30.Text = ""
            TextBox31.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            PictureBox3.Image = Nothing

        ElseIf result = DialogResult.Cancel Then
            PictureBox3.Image = Nothing
        End If
    End Sub
End Class