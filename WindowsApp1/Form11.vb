Imports System.Drawing.Imaging
Imports IDAutomation.Windows.Forms.LinearBarCode
Imports System.Drawing.Printing
Imports System.Configuration
Imports GenCode128
Imports MySql.Data.MySqlClient
Public Class Form11
    Private WithEvents pdPrint As PrintDocument
    Private PrintDocType As String = "Barcode"
    Private StrPrinterName As String = "Canon MP280 series"
    Public old_barcode_id As String

    Dim pbImage2 As New PictureBox
    Private bmp As Bitmap
    Dim document_type As String = "10"
    Dim barcode_id As String



    Dim prtdoc As New PrintDocument
    Dim strDefaultPrinter As String = prtdoc.PrinterSettings.PrinterName


    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Clear text box
        Text_Name1.Text = ""
        Text_Name2.Text = ""
        Text_Name3.Text = ""
        Text_Name1.Text = ""
        Text_Name2.Text = ""
        Text_Name3.Text = ""
        Text_ADV.Text = ""
        Text_Year.Text = ""
        Text_NAMEP.Text = ""
        Text_P3.Text = ""
        Text_P2.Text = ""
        Text_P1.Text = ""
        Text_ID1.Text = ""
        Text_ID2.Text = ""
        Text_ID3.Text = ""
        Text_A.Text = ""
        TextBox1.Text = ""
        TextBox18.Text = ""
        TextBox2.Text = ""
        TextBox23.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""

    End Sub

    Public Sub Print_Document()
        bmp = New Bitmap(Panel1.Width, Panel1.Height)
        Dim G As Graphics = Graphics.FromImage(bmp)

        Panel1.DrawToBitmap(bmp, Panel1.ClientRectangle)
        G.Dispose()

        PrintDocument1.DefaultPageSettings.PaperSize = New PaperSize("210 x 297 mm", 761, 1084)
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


        Dim MyDate As String = DateTime.Now.ToString("HHmmddMMyy")
        'date_time_now = DateTime.Now.ToString
        barcode_id = Text_ID1.Text.ToString() & MyDate
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

        PictureBox8.Image = Image.FromFile(Application.StartupPath & "\" & "SavedBarcode.Jpeg")

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
            ' INSERT INTO `Project_std`(`ID1`, `ID2`, `ID3`, `NAME1`, `NAME2`, `NAME3`, `PHONE1`, `PHONE2`, `PHONE3`, `PROJECTNAME`, `PROJECTADVISOR`, `SEMESTER`, `YEARS`, `STATUS`, `TYPE`, `Barcode`, `DATE`) VALUES ([value-1],[value-2],[value-3],[value-4],[value-5],[value-6],[value-7],[value-8],[value-9],[value-10],[value-11],[value-12],[value-13],[value-14],[value-15],[value-16],[value-17])
            Query_CMD = "INSERT INTO `Project_std`(`ID1`,`ID2`,`ID3`,`NAME1`,`NAME2`,`NAME3`,`PHONE1`,`PHONE2`,`PHONE3`,`PROJECTNAME`,`PROJECTADVISOR`, `STATUS`,`TYPE`,`Barcode`) VALUES ('" & Text_ID1.Text & "','" & Text_ID2.Text & "','" & Text_ID3.Text & "','" & Text_Name1.Text & "','" & Text_Name2.Text & "','" & Text_Name3.Text & "','" & Text_P1.Text & "','" & Text_P2.Text & "','" & Text_P3.Text & "','" & Text_NAMEP.Text & "','" & Text_A.Text & "','ส่งเอกสาร','" & document_type & "','" & barcode_id & "')"

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
        ' Dim Query As String = "INSERT INTO `inform_std` (STD_ID,STD_NAME,STD_SUBJECT,STD_IDSub,STD_ADVISOR,STD_LECTURER,DATE,STD_STATUS,STD_TYPE) VALUES ('B12344','ปอร์ด','ไมโครเวฟ','123456','John','Jame','2019-4-22','...','12')"
        ' Dim  As String = "SET character_set_connection=utf8"
        'TextBox1.Text = TextBox1.Text & "Qury Text:" & Query & vbCrLf
        ' Query = " (eid,name,surname,age) values ('" & TextBox_Eid.Text & "','" & TextBox_Name.Text & "','" & TextBox_SName.Text & "','" & TextBox_Age.Text & "')"
        Dim NewBarcode As IDAutomation.Windows.Forms.LinearBarCode.Barcode = New Barcode()

        Dim MyDate As String = DateTime.Now.ToString("HHmmddMMyy")
        'date_time_now = DateTime.Now.ToString
        barcode_id = Text_ID1.Text.ToString() & MyDate
        'MsgBox(barcode_id)

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

        PictureBox8.Image = Image.FromFile(Application.StartupPath & "\" & "SavedBarcode.Jpeg")

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
            MsgBox("SQL Database Connect OK!")

            Dim Query_CMD As String
            'UPDATE `Project_std` SET `ID1`=[value-1],`ID2`=[value-2],`ID3`=[value-3],`NAME1`=[value-4],`NAME2`=[value-5],`NAME3`=[value-6],`PHONE1`=[value-7],`PHONE2`=[value-8],`PHONE3`=[value-9],`PROJECTNAME`=[value-10],`PROJECTADVISOR`=[value-11],`SEMESTER`=[value-12],`YEARS`=[value-13],`STATUS`=[value-14],`TYPE`=[value-15],`Barcode`=[value-16],`DATE`=[value-17] WHERE 1
            Query_CMD = "UPDATE `Project_std` SET `ID1`='" & Text_ID1.Text & "',`ID2`='" & Text_ID2.Text & "',`ID3`='" & Text_ID3.Text & "',`NAME1`='" & Text_Name1.Text & "',`NAME2`='" & Text_Name2.Text & "',`NAME3`='" & Text_Name3.Text & "',`PHONE1`='" & Text_P1.Text & "',`PHONE2`='" & Text_P2.Text & "',`PHONE3`='" & Text_P3.Text & "',`PROJECTNAME`='" & Text_NAMEP.Text & "',`PROJECTADVISOR`='" & Text_A.Text & "',`YEARS`='" & Text_Year.Text & "' WHERE `Barcode`='" & old_barcode_id.ToString & "'"
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

            Text_Name1.Text = ""
            Text_Name2.Text = ""
            Text_Name3.Text = ""
            Text_Name1.Text = ""
            Text_Name2.Text = ""
            Text_Name3.Text = ""
            Text_ADV.Text = ""
            Text_Year.Text = ""
            Text_NAMEP.Text = ""
            Text_P3.Text = ""
            Text_P2.Text = ""
            Text_P1.Text = ""
            Text_ID1.Text = ""
            Text_ID2.Text = ""
            Text_ID3.Text = ""
            Text_A.Text = ""
            TextBox1.Text = ""
            TextBox18.Text = ""
            TextBox2.Text = ""
            TextBox23.Text = ""
            TextBox3.Text = ""
            TextBox5.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox3.Text = ""
            PictureBox1.Image = Nothing


        ElseIf result = DialogResult.No Then
            Text_Name1.Text = ""
            Text_Name2.Text = ""
            Text_Name3.Text = ""
            Text_Name1.Text = ""
            Text_Name2.Text = ""
            Text_Name3.Text = ""
            Text_ADV.Text = ""
            Text_Year.Text = ""
            Text_NAMEP.Text = ""
            Text_P3.Text = ""
            Text_P2.Text = ""
            Text_P1.Text = ""
            Text_ID1.Text = ""
            Text_ID2.Text = ""
            Text_ID3.Text = ""
            Text_A.Text = ""
            TextBox1.Text = ""
            TextBox18.Text = ""
            TextBox2.Text = ""
            TextBox23.Text = ""
            TextBox3.Text = ""
            TextBox5.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox3.Text = ""

            PictureBox8.Image = Nothing

        ElseIf result = DialogResult.Cancel Then
            PictureBox8.Image = Nothing
        End If
    End Sub
    Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        Insert_Edit.Show()
    End Sub
    Private Sub Button_Edit_Click(sender As Object, e As EventArgs) Handles Button_Edit.Click
        Update_SQL()
    End Sub

End Class