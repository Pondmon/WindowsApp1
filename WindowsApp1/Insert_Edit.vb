Imports MySql.Data.MySqlClient


Public Class Insert_Edit
    Public form_ID_callback As String

    Private Sub Insert_Edit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Label1.Text = form_ID_callback
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

            If dt.Rows.Count > 0 Then  ' check data form table 1 of SQL

                '---- Call back 
                If form_ID_callback = "Form_2" Then
                    Form2.Text_Name.Text = dt.Rows(0)("STD_NAME")
                    Form2.Text_ID.Text = dt.Rows(0)("STD_ID")
                    Form2.Text_IDSUB.Text = dt.Rows(0)("STD_IDSub")
                    Form2.Text_SUB.Text = dt.Rows(0)("STD_SUBJECT")
                    Form2.ComboBox4.Text = dt.Rows(0)("STD_LECTURER")
                    Form2.Text_INS.Text = dt.Rows(0)("STD_Faculty")
                    Form2.Text_School.Text = dt.Rows(0)("STD_School")

                    Form2.old_barcode_id = dt.Rows(0)("STD_Barcode")

                ElseIf form_ID_callback = "Form_3" Then
                    Form3.Text_Name.Text = dt.Rows(0)("STD_NAME")
                    Form3.Text_ID.Text = dt.Rows(0)("STD_ID")
                    Form3.Text_IDSUB.Text = dt.Rows(0)("STD_IDSub")
                    Form3.Text_SUB.Text = dt.Rows(0)("STD_SUBJECT")
                    Form3.ComboBox4.Text = dt.Rows(0)("STD_LECTURER")
                    Form3.Text_ADV.Text = dt.Rows(0)("STD_ADVISOR")
                    Form3.Text_INS.Text = dt.Rows(0)("STD_Faculty")
                    Form3.Text_School.Text = dt.Rows(0)("STD_School")
                    Form3.Text_R.Text = dt.Rows(0)("STD_Reason")

                    Form3.old_barcode_id = dt.Rows(0)("STD_Barcode")

                ElseIf form_ID_callback = "Form_4" Then
                    Form4.Text_NAME.Text = dt.Rows(0)("STD_NAME")
                    Form4.Text_ID.Text = dt.Rows(0)("STD_ID")
                    Form4.Text_IDSUB1.Text = dt.Rows(0)("STD_IDSub")
                    Form4.Text_SUB1.Text = dt.Rows(0)("STD_SUBJECT")
                    Form4.Text_ADV.Text = dt.Rows(0)("STD_LECTURER")
                    Form4.Text_IDSUB2.Text = dt.Rows(0)("STD_IDSub2")
                    Form4.Text_IDSUB3.Text = dt.Rows(0)("STD_IDSub3")
                    Form4.Text_IDSUB4.Text = dt.Rows(0)("STD_IDSub4")
                    Form4.Text_SUB2.Text = dt.Rows(0)("STD_SUBJECT2")
                    Form4.Text_SUB4.Text = dt.Rows(0)("STD_SUBJECT4")
                    Form4.Text_SUB3.Text = dt.Rows(0)("STD_SUBJECT3")
                    Form4.Text_INS.Text = dt.Rows(0)("STD_Faculty")
                    Form4.Text_School.Text = dt.Rows(0)("STD_School")
                    Form4.Text_Pro.Text = dt.Rows(0)("STD_Professor")


                    Form4.old_barcode_id = dt.Rows(0)("STD_Barcode")

                ElseIf form_ID_callback = "Form_5" Then
                    Form5.Text_NAME.Text = dt.Rows(0)("STD_NAME")
                    Form5.Text_ID.Text = dt.Rows(0)("STD_ID")
                    Form5.Text_IDSUB1.Text = dt.Rows(0)("STD_IDSub")
                    Form5.Text_SUB1.Text = dt.Rows(0)("STD_SUBJECT")
                    Form5.ComboBox3.Text = dt.Rows(0)("STD_ADVISOR")
                    Form5.Text_INS.Text = dt.Rows(0)("STD_Faculty")
                    Form5.Text_School.Text = dt.Rows(0)("STD_School")

                    Form5.old_barcode_id = dt.Rows(0)("STD_Barcode")

                ElseIf form_ID_callback = "Form_6" Then
                    Form6.Text_Name.Text = dt.Rows(0)("STD_NAME")
                    Form6.Text_ID.Text = dt.Rows(0)("STD_ID")
                    Form6.Text_IDSUB.Text = dt.Rows(0)("STD_IDSub")
                    Form6.Text_SUB.Text = dt.Rows(0)("STD_SUBJECT")
                    Form6.Text_S.Text = dt.Rows(0)("STD_LECTURER")
                    Form6.ComboBox3.Text = dt.Rows(0)("STD_ADVISOR")
                    Form6.Text_INS.Text = dt.Rows(0)("STD_Faculty")
                    Form6.Text_School.Text = dt.Rows(0)("STD_School")

                    Form3.old_barcode_id = dt.Rows(0)("STD_Barcode")

                ElseIf form_ID_callback = "Form_7" Then
                    Form7.Text_Name.Text = dt.Rows(0)("STD_NAME")
                    Form7.Text_ID.Text = dt.Rows(0)("STD_ID")
                    Form7.Text_IDSUB.Text = dt.Rows(0)("STD_IDSub")
                    Form7.Text_SUB1.Text = dt.Rows(0)("STD_SUBJECT")
                    Form7.TextIDSUB2.Text = dt.Rows(0)("STD_IDSub2")
                    Form7.Text_SUB2.Text = dt.Rows(0)("STD_SUBJECT2")
                    Form7.ComboBox4.Text = dt.Rows(0)("STD_LECTURER")
                    Form7.Text_INS.Text = dt.Rows(0)("STD_Faculty")
                    Form7.Text_School.Text = dt.Rows(0)("STD_School")

                    Form7.old_barcode_id = dt.Rows(0)("STD_Barcode")

                ElseIf form_ID_callback = "Form_8" Then

                    Form8.Text_Name.Text = dt.Rows(0)("STD_NAME")
                    Form8.Text_ID.Text = dt.Rows(0)("STD_ID")
                    Form8.ComboBox3.Text = dt.Rows(0)("STD_ADVISOR")
                    Form8.Text_INS.Text = dt.Rows(0)("STD_Faculty")
                    Form8.Text_School.Text = dt.Rows(0)("STD_School")
                    Form8.Text_Pro.Text = dt.Rows(0)("STD_Professor")

                    Form8.old_barcode_id = dt.Rows(0)("STD_Barcode")

                ElseIf form_ID_callback = "Form_9" Then
                    Form9.Text_Name.Text = dt.Rows(0)("STD_NAME")
                    Form9.Text_ID.Text = dt.Rows(0)("STD_ID")
                    Form9.ComboBox3.Text = dt.Rows(0)("STD_ADVISOR")
                    Form9.Text_Pro.Text = dt.Rows(0)("STD_Professor")
                    Form9.Text_INS.Text = dt.Rows(0)("STD_Faculty")
                    Form9.Text_School.Text = dt.Rows(0)("STD_School")

                    Form9.old_barcode_id = dt.Rows(0)("STD_Barcode")
                End If

            Else

                ' Read second table
                Query_CMD = "SELECT * FROM `Project_std` WHERE `Barcode` = '" & Me.TEXT_INSERT_BARCODE.Text & "'"
                'Me.Text_information.Text = Text_information.Text + Query_CMD

                dtAdapter = New MySqlDataAdapter(Query_CMD, conn)
                dtAdapter.Fill(dt)

                If form_ID_callback = "Form_10" Then
                    Form10.Text_Name1.Text = dt.Rows(0)("NAME1")
                    Form10.Text_Name2.Text = dt.Rows(0)("NAME2")
                    Form10.Text_Name3.Text = dt.Rows(0)("NAME3")
                    Form10.Text_ID1.Text = dt.Rows(0)("ID1")
                    Form10.Text_ID2.Text = dt.Rows(0)("ID2")
                    Form10.Text_ID3.Text = dt.Rows(0)("ID3")
                    Form10.Text_NAMEP.Text = dt.Rows(0)("PROJECTNAME")
                    Form10.Text_P1.Text = dt.Rows(0)("PHONE1")
                    Form10.Text_P2.Text = dt.Rows(0)("PHONE2")
                    Form10.Text_P3.Text = dt.Rows(0)("PHONE3")
                    Form10.ComboBox4.Text = dt.Rows(0)("PROJECTADVISOR")

                    Form10.old_barcode_id = dt.Rows(0)("Barcode")

                ElseIf form_ID_callback = "Form_11" Then
                    Form11.Text_Name1.Text = dt.Rows(0)("NAME1")
                    Form11.Text_Name2.Text = dt.Rows(0)("NAME2")
                    Form11.Text_Name3.Text = dt.Rows(0)("NAME3")
                    Form11.Text_ID1.Text = dt.Rows(0)("ID1")
                    Form11.Text_ID2.Text = dt.Rows(0)("ID2")
                    Form11.Text_ID3.Text = dt.Rows(0)("ID3")
                    Form11.Text_NAMEP.Text = dt.Rows(0)("PROJECTNAME")
                    Form11.Text_P1.Text = dt.Rows(0)("PHONE1")
                    Form11.Text_P2.Text = dt.Rows(0)("PHONE2")
                    Form11.Text_P3.Text = dt.Rows(0)("PHONE3")
                    Form11.ComboBox4.Text = dt.Rows(0)("PROJECTADVISOR")

                    Form11.old_barcode_id = dt.Rows(0)("Barcode")

                End If
            End If

            conn.Close()
            conn = Nothing

            Me.Close()

        Catch ex As MySql.Data.MySqlClient.MySqlException
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try

        form_ID_callback = ""
    End Sub
End Class