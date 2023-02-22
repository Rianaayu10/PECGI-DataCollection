Imports System.Data.SqlClient
Imports System.IO

Public Class FormConfig
#Region "INITIAL"

    Dim NewEnryption As New clsDESEncryption("TOS")
    Dim lb_Load As Boolean = True

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

#End Region

#Region "EVENTS"
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnTestCon_Click(sender As Object, e As EventArgs) Handles btnTestCon.Click
        up_TestConnection()
    End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        If MsgBox("Are you sure want to apply this setting?", MsgBoxStyle.YesNo + vbDefaultButton2 + MsgBoxStyle.Information, "Apply Setting Confirmation") = MsgBoxResult.Yes Then
            up_ApplySetting()
        End If
    End Sub

    Private Sub CheckWindowsNT_CheckedChanged(sender As Object, e As EventArgs) Handles CheckWindowsNT.CheckedChanged
        If CheckWindowsNT.Checked = True Then
            TxtUserId.Enabled = False
            TxtPassword.Enabled = False
            TxtDbTimeOut.Enabled = False
            TxtCmdTimeOut.Enabled = False
        Else
            TxtUserId.Enabled = True
            TxtPassword.Enabled = True
            TxtDbTimeOut.Enabled = True
            TxtCmdTimeOut.Enabled = True
        End If
    End Sub

    Private Sub cboMachineType_TextChanged(sender As Object, e As EventArgs) Handles cboMachineType.TextChanged
        up_LoadData()
    End Sub

    Private Sub FormConfig_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            '# Setting...
            Dim ls_path As String

            ls_path = My.Application.Info.DirectoryPath & "\config.xml"

            If My.Computer.FileSystem.FileExists(ls_path) = False Then
                Dim fs1 As FileStream = New FileStream(ls_path, FileMode.Create, FileAccess.Write)
                Dim s1 As StreamWriter = New StreamWriter(fs1)

                s1.Close()
                fs1.Close()
                MsgBox("Config file is not found!", MsgBoxStyle.Information, "Info")

                ActiveControl = TxtServerName
                Exit Sub
            End If

            Call up_AppSettingLoad(ls_path)
            Call up_Combo()
            Call up_ClearFTP()

            ActiveControl = TxtServerName
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "PROCEDURE"
    Private Sub up_TestConnection()
        Dim ls_con As String
        Dim con As New SqlConnection

        Try
            If CheckWindowsNT.Checked = False Then
                ls_con = "Data Source=" & TxtServerName.Text & ";Initial Catalog=" & TxtDatabase.Text & ";User ID=" & TxtUserId.Text & ";pwd=" & TxtPassword.Text & ""
            Else
                ls_con = "Data Source=" & TxtServerName.Text & ";Initial Catalog=" & TxtDatabase.Text & ";Integrated Security=True"
            End If
            con.ConnectionString = ls_con
            con.Open()

            If con.State = ConnectionState.Open Then
                MsgBox("Test Connection EZ Runner Succeeded !", MsgBoxStyle.Information, "Info")
            End If

            con.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub up_ApplySetting()
        Dim docXML = New XDocument()

        Dim crtElement = New XElement("Connection")
        Dim SMType As String = ""

        '01. Connection Database Setting
        Dim ASDB = New XElement("EZDB")
        ASDB.Add(New XElement("ServerName", NewEnryption.EncryptData(TxtServerName.Text)))
        ASDB.Add(New XElement("Database", NewEnryption.EncryptData(TxtDatabase.Text)))
        ASDB.Add(New XElement("WinMode", NewEnryption.EncryptData(IIf(CheckWindowsNT.Checked = True, "", "mixed"))))
        ASDB.Add(New XElement("UserID", NewEnryption.EncryptData(TxtUserId.Text)))
        ASDB.Add(New XElement("Password", NewEnryption.EncryptData(TxtPassword.Text)))
        ASDB.Add(New XElement("DatabaseTimeOut", NewEnryption.EncryptData(TxtDbTimeOut.Text)))
        ASDB.Add(New XElement("CommandTimeOut", NewEnryption.EncryptData(TxtCmdTimeOut.Text)))
        crtElement.Add(ASDB)


        Dim PATHBA = New XElement("PATHBA")
        PATHBA.Add(New XElement("ServerPath", NewEnryption.EncryptData(txtBobbinArrangingPath.Text)))
        PATHBA.Add(New XElement("LocalPath", NewEnryption.EncryptData(txtBobbinArrangingLocalPath.Text)))
        PATHBA.Add(New XElement("User", NewEnryption.EncryptData(txtBobbinArrangingUser.Text)))
        PATHBA.Add(New XElement("Password", NewEnryption.EncryptData(txtBobbinArrangingPassword.Text)))
        crtElement.Add(PATHBA)

        docXML.Add(crtElement)
        docXML.Save("config.xml")

        up_InsUpdData()

        MsgBox("Change Apply Setting Succeeded." & Chr(13) & "You need to restart your application.", MsgBoxStyle.Information, "Apply Setting")
    End Sub

    Private Sub up_AppSettingLoad(ByVal ls_path As String)
        Try
            'check if file myxml.xml is existing
            If (IO.File.Exists(ls_path)) Then

                'Check XML file is empty or not
                If Trim(IO.File.ReadAllText(ls_path).Length) = 0 Then Exit Sub

                'Load XML file
                Dim document = XDocument.Load(ls_path)

                ''''''''''''''''''''''''''''''''''01. Load Setting ASDB to Screen''''''''''''''''''''''''''''''''''
                Dim ASDB = document.Descendants("EZDB").FirstOrDefault()

                If Not IsNothing(ASDB) Then
                    If Not IsNothing(ASDB.Element("ServerName")) Then TxtServerName.Text = NewEnryption.DecryptData(ASDB.Element("ServerName").Value)
                    If Not IsNothing(ASDB.Element("Database")) Then TxtDatabase.Text = NewEnryption.DecryptData(ASDB.Element("Database").Value)
                    If Not IsNothing(ASDB.Element("UserID")) Then TxtUserId.Text = NewEnryption.DecryptData(ASDB.Element("UserID").Value)
                    If Not IsNothing(ASDB.Element("Password")) Then TxtPassword.Text = NewEnryption.DecryptData(ASDB.Element("Password").Value)
                    If Not IsNothing(ASDB.Element("DatabaseTimeOut")) Then TxtDbTimeOut.Text = NewEnryption.DecryptData(ASDB.Element("DatabaseTimeOut").Value)
                    If Not IsNothing(ASDB.Element("CommandTimeOut")) Then TxtCmdTimeOut.Text = NewEnryption.DecryptData(ASDB.Element("CommandTimeOut").Value)
                    If Not IsNothing(ASDB.Element("WinMode")) Then
                        If NewEnryption.DecryptData(ASDB.Element("WinMode").Value).ToLower <> "mixed" Then
                            CheckWindowsNT.Checked = True
                        Else
                            CheckWindowsNT.Checked = False
                        End If
                    End If
                End If
                ''''''''''''''''''''''''''''''''''END''''''''''''''''''''''''''''''''''


                Dim PATHBA = document.Descendants("PATHBA").FirstOrDefault()

                If Not IsNothing(PATHBA) Then
                    If Not IsNothing(PATHBA.Element("ServerPath")) Then txtBobbinArrangingPath.Text = NewEnryption.DecryptData(PATHBA.Element("ServerPath").Value)
                    If Not IsNothing(PATHBA.Element("LocalPath")) Then txtBobbinArrangingLocalPath.Text = NewEnryption.DecryptData(PATHBA.Element("LocalPath").Value)
                    If Not IsNothing(PATHBA.Element("User")) Then txtBobbinArrangingUser.Text = NewEnryption.DecryptData(PATHBA.Element("User").Value)
                    If Not IsNothing(PATHBA.Element("Password")) Then txtBobbinArrangingPassword.Text = NewEnryption.DecryptData(PATHBA.Element("Password").Value)
                End If

            Else
                MessageBox.Show("Config File is not found!")
            End If
            lb_Load = False
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Exclamation, "Warning")
        End Try
    End Sub

    Private Sub up_Combo()
        With cboMachineType
            .DataMode = C1.Win.C1List.DataModeEnum.AddItem
            .ColumnHeaders = False
            .ClearItems()
            .AddItem("CaseCarbon X-01")
            .AddItem("CaseCarbon X-02")
            .AddItem("CaseCarbon X-03")
            .AddItem("CaseCarbon X-04")
            .AddItem("CaseCarbon X-05")
            .ExtendRightColumn = True
            .DropDownWidth = .Width
            .HScrollBar.Height = 0
            .SelectedIndex = -1
        End With
    End Sub

    Private Sub up_ClearFTP()
        cboMachineType.SelectedIndex = -1
        txtBobbinArrangingPath.Text = ""
        txtBobbinArrangingLocalPath.Text = ""
        txtBobbinArrangingUser.Text = ""
        txtBobbinArrangingPassword.Text = ""
    End Sub

    Private Sub up_LoadData()
        Dim sql As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As New DataSet

        Try

            Using connection As New SqlConnection(ConStr)
                connection.Open()

                sql = "SP_FTPLoad_Data"

                cmd = New SqlCommand(sql, connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("Type", Trim(cboMachineType.Text))

                da = New SqlDataAdapter(cmd)
                da.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    txtBobbinArrangingPath.Text = ds.Tables(0).Rows(0)("Server_Path").ToString.Trim
                    txtBobbinArrangingLocalPath.Text = ds.Tables(0).Rows(0)("Local_Path").ToString.Trim
                    txtBobbinArrangingUser.Text = ds.Tables(0).Rows(0)("user_FTP").ToString.Trim
                    txtBobbinArrangingPassword.Text = ds.Tables(0).Rows(0)("Password_FTP").ToString.Trim
                Else
                    txtBobbinArrangingPath.Text = ""
                    txtBobbinArrangingLocalPath.Text = ""
                    txtBobbinArrangingUser.Text = ""
                    txtBobbinArrangingPassword.Text = ""
                End If

            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub up_InsUpdData()
        Dim sql As String = ""
        Try
            Using connection As New SqlConnection(ConStr)
                connection.Open()

                sql = "SP_FTPInsUpd_Data"

                Dim cmd As New SqlCommand(sql, connection)
                cmd.CommandType = CommandType.StoredProcedure

                With cmd.Parameters
                    .AddWithValue("Type", Trim(cboMachineType.Text))
                    .AddWithValue("ServerPath", Trim(txtBobbinArrangingPath.Text))
                    .AddWithValue("LocalPath", Trim(txtBobbinArrangingLocalPath.Text))
                    .AddWithValue("UserFTP", Trim(txtBobbinArrangingUser.Text))
                    .AddWithValue("Password", Trim(txtBobbinArrangingPassword.Text))
                End With

                cmd.ExecuteNonQuery()

            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

End Class