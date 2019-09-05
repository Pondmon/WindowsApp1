Imports MySql.Data.MySqlClient


Public Class Insert_Edit

    Private Sub Insert_Edit_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
            'SELECT * FROM `inform_std` WHERE `STD_ID` = '12314'
            Query_CMD = "SELECT * FROM `inform_std` WHERE `STD_Barcode` = '" & Me.TEXT_INSERT_BARCODE.Text & "'"
            'Me.TextBox1.Text = TextBox1.Text + Query_CMD

            Dim dtAdapter As MySqlDataAdapter
            Dim dt As New DataTable

            dtAdapter = New MySqlDataAdapter(Query_CMD, conn)
            dtAdapter.Fill(dt)

            If dt.Rows.Count > 0 Then
                'Me.TextBox1.Text = TextBox1.Text + dt.Rows(0)("STD_ID") & vbCrLf
                'Me.TextBox1.Text = TextBox1.Text + dt.Rows(0)("STD_NAME") & vbCrLf
                'Me.TextBox1.Text = TextBox1.Text + dt.Rows(0)("STD_STATUS") & vbCrLf

                Form6.Text_Name.Text = dt.Rows(0)("STD_NAME")
                Form6.Text_ID.Text = dt.Rows(0)("STD_ID")
                Form6.Text_IDSUB.Text = dt.Rows(0)("STD_IDSub")
                Form6.Text_SUB.Text = dt.Rows(0)("STD_SUBJECT")
                Form6.Text_S.Text = dt.Rows(0)("STD_LECTURER")

                Form6.old_barcode_id = dt.Rows(0)("STD_Barcode")

            End If

            conn.Close()
            conn = Nothing

            Me.Close()

        Catch ex As MySql.Data.MySqlClient.MySqlException
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try
    End Sub


End Class