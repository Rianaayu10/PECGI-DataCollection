Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Threading
Imports C1.Win.C1FlexGrid

Public Class FormSchedulerCSV_Packing_Quality
#Region "INITIAL"

    Private cConfig As clsConfig

    Const ConnectionErrorMsg As String = "A network-related or instance-specific error occurred while establishing a connection to SQL Server"
    Const TransportErrorMsg As String = "A transport-level error has occurred"

    Public Sub New()
        InitializeComponent()
    End Sub

#End Region

#Region "DECLARATION"

    Dim ShrinkCutting1 As Byte = 1
    Dim WeightBatch1 As Byte = 2
    Dim Heater1 As Byte = 3
    Dim Taping1 As Byte = 4
    Dim WeightBox1 As Byte = 5
    Dim ShrinkCutting2 As Byte = 6
    Dim WeightBatch2 As Byte = 7
    Dim Heater2 As Byte = 8
    Dim Taping2 As Byte = 9
    Dim WeightBox2 As Byte = 10
    Dim PillowProcess As Byte = 11

    Dim col_ProcessName As Integer = 0
    Dim Col_ProcessType As Byte = 1
    Dim Col_ProcessStatus As Byte = 2
    Dim Col_LastProcess As Byte = 3
    Dim Col_NextProcess As Byte = 4
    Dim Col_ErrorMessage As Byte = 5

    Dim col_Count As Integer = 6

    Dim FilesList As New List(Of String)

    Dim Thd_ShrinkCutting1 As SchedulerSetting
    Dim Thd_WeightBatch1 As SchedulerSetting
    Dim Thd_Heater1 As SchedulerSetting
    Dim Thd_Taping1 As SchedulerSetting
    Dim Thd_WeightBox1 As SchedulerSetting
    Dim Thd_ShrinkCutting2 As SchedulerSetting
    Dim Thd_WeightBatch2 As SchedulerSetting
    Dim Thd_Heater2 As SchedulerSetting
    Dim Thd_Taping2 As SchedulerSetting
    Dim Thd_WeightBox2 As SchedulerSetting
    Dim Thd_PillowProcess As SchedulerSetting

    Dim m_Finish As Boolean

    Dim log As clsLog

    Private Enum ExcelCols
        col_Date = 0
        col_Time = 1
        col_C = 2
        col_D = 3
        col_E = 4
        col_F = 5
        col_G = 6
        col_H = 7
        col_I = 8
        col_J = 9
        col_K = 10
        col_L = 11
        col_M = 12
        col_N = 13
        col_O = 14
        col_P = 15
        col_Q = 16
        col_R = 17
        col_S = 18
        col_T = 19
        col_U = 20
    End Enum

    Public Structure SchedulerSetting
        Public Name As String
        Public ScheduleThd As Thread
        Public Lock As Object
        Public EndSchedule As Boolean
        Public DelayTime As Double
        Public Status As String
    End Structure

#End Region

#Region "EVENTS"
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnConfig_Click(sender As Object, e As EventArgs) Handles btnConfig.Click
        FormConfig.ShowDialog()

        Application.DoEvents()
        cConfig = New clsConfig

        gs_Server = cConfig.Server
        gs_Database = cConfig.Database
        gs_User = cConfig.User
        gs_Password = cConfig.Password
        ConStr = cConfig.ConnectionString

        If gs_Server = "" Then
            stpStatus.Text = "Database : No connection to server"
        Else
            stpStatus.Text = gs_Server & "." & gs_Database
        End If

        up_LoadData()
    End Sub

    Private Sub btnManual_Click(sender As Object, e As EventArgs) Handles btnManual.Click
        txtMsg.Text = ""
        up_Process()
    End Sub

    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        txtMsg.Text = ""

        up_TimeStop()

        Do Until Thd_ShrinkCutting1.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_ShrinkCutting1.Status = "iddle" Then
                Thd_ShrinkCutting1.ScheduleThd = Nothing
            End If
            If Thd_ShrinkCutting1.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        Do Until Thd_WeightBatch1.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_WeightBatch1.Status = "iddle" Then
                Thd_WeightBatch1.ScheduleThd = Nothing
            End If
            If Thd_WeightBatch1.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        Do Until Thd_Heater1.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_Heater1.Status = "iddle" Then
                Thd_Heater1.ScheduleThd = Nothing
            End If
            If Thd_Heater1.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        Do Until Thd_WeightBox1.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_WeightBox1.Status = "iddle" Then
                Thd_WeightBox1.ScheduleThd = Nothing
            End If
            If Thd_WeightBox1.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        Do Until Thd_Taping1.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_Taping1.Status = "iddle" Then
                Thd_Taping1.ScheduleThd = Nothing
            End If
            If Thd_Taping1.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        Do Until Thd_ShrinkCutting2.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_ShrinkCutting2.Status = "iddle" Then
                Thd_ShrinkCutting2.ScheduleThd = Nothing
            End If
            If Thd_ShrinkCutting2.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(200)
        Loop

        Do Until Thd_WeightBatch2.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_WeightBatch2.Status = "iddle" Then
                Thd_WeightBatch2.ScheduleThd = Nothing
            End If
            If Thd_WeightBatch2.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(200)
        Loop

        Do Until Thd_Heater2.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_Heater2.Status = "iddle" Then
                Thd_Heater2.ScheduleThd = Nothing
            End If
            If Thd_Heater2.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(200)
        Loop

        Do Until Thd_WeightBox2.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_WeightBox2.Status = "iddle" Then
                Thd_WeightBox2.ScheduleThd = Nothing
            End If
            If Thd_WeightBox2.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(200)
        Loop

        Do Until Thd_Taping2.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_Taping2.Status = "iddle" Then
                Thd_Taping2.ScheduleThd = Nothing
            End If
            If Thd_Taping2.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(200)
        Loop

        Do Until Thd_PillowProcess.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_PillowProcess.Status = "iddle" Then
                Thd_PillowProcess.ScheduleThd = Nothing
            End If
            If Thd_PillowProcess.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(200)
        Loop

        btnStart.Enabled = True
        btnManual.Enabled = True
        btnConfig.Enabled = True
        btnClose.Enabled = True

        up_GridHeader()
        up_GridLoad()
        up_LoadData()
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        txtMsg.Text = ""

        up_TimeStart()

        btnStart.Enabled = False
        btnStop.Enabled = True
        btnManual.Enabled = False
        btnConfig.Enabled = False
        btnClose.Enabled = False
    End Sub

    Private Sub timerCurr_Tick(sender As Object, e As EventArgs) Handles timerCurr.Tick
        lblCurrTime.Text = Format(Now, "HH:mm:ss")
        lblCurrDate.Text = Format(Now, "dddd , dd MMM yyyy")
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        ShowInTaskbar = True
        Me.WindowState = FormWindowState.Normal
        NotifyIcon1.Visible = False
    End Sub

    Private Sub FormSchedulerCSV_Packing_Quality_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtMsg.Text = ""
        Dim strVersion As String = Me.ProductVersion

        lblVersion.Text = strVersion

        timerCurr.Enabled = True

        Application.DoEvents()
        cConfig = New clsConfig

        If gs_Server = "" Then
            stpStatus.Text = "Database : No connection to server"
        Else
            stpStatus.Text = gs_Server & "." & gs_Database
        End If

        up_GridHeader()
        up_GridLoad()
        up_LoadData()
    End Sub

    Private Sub FormSchedulerCSV_Packing_Quality_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                'NotifyIcon1.Icon = SystemIcons.Application
                NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
                NotifyIcon1.BalloonTipTitle = "RELAY PLC CSV QUALITY (PACKING)"
                NotifyIcon1.BalloonTipText = "Move to system tray"
                NotifyIcon1.ShowBalloonTip(50)
                'Me.Hide()
                ShowInTaskbar = False
            End If
        End If
    End Sub
#End Region

#Region "PROCEDURES"

    Private Sub up_GridHeader()
        With grid
            .Rows.Fixed = 1
            .Rows.Count = .Rows.Fixed
            .AllowEditing = False
            .Cols.Count = col_Count
            .Rows(0).AllowMerging = True
            .Rows(0).Height = 30
            .Styles.Normal.WordWrap = True

            .Item(0, col_ProcessName) = "PROCESS NAME"
            .Item(0, Col_ProcessType) = "PROCESS TYPE"
            .Item(0, Col_ProcessStatus) = "PROCESS STATUS"
            .Item(0, Col_LastProcess) = "LAST PROCESS"
            .Item(0, Col_NextProcess) = "NEXT PROCESS"
            .Item(0, Col_ErrorMessage) = "ERROR MESSAGE"

            .Cols(col_ProcessName).Width = 90
            .Cols(Col_ProcessType).Width = 220
            .Cols(Col_ProcessStatus).Width = 100
            .Cols(Col_LastProcess).Width = 120
            .Cols(Col_NextProcess).Width = 120
            .Cols(Col_ErrorMessage).Width = 300

            .Rows(0).StyleNew.TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.CenterCenter

            up_FillBackColourGrid(grid, col_Count - 1, 0)

        End With
    End Sub

    Private Sub up_GridLoad()
        Dim sql As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As New DataSet

        Dim li_Row As Integer = 0

        Try

            Using connection As New SqlConnection(ConStr)
                connection.Open()

                sql = "SP_FTPLoad_Grid"

                cmd = New SqlCommand(sql, connection)
                cmd.CommandType = CommandType.StoredProcedure

                da = New SqlDataAdapter(cmd)
                da.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then

                    li_Row = 1
                    For i As Integer = 0 To ds.Tables(0).Rows.Count - 1

                        grid.AddItem("")
                        grid.Item(li_Row, col_ProcessName) = ds.Tables(0).Rows(i)("ProcessName").ToString.Trim
                        grid.Item(li_Row, Col_ProcessType) = ds.Tables(0).Rows(i)("ProcessType").ToString.Trim
                        grid.Item(li_Row, Col_ProcessStatus) = ds.Tables(0).Rows(i)("ProcessStatus").ToString.Trim
                        grid.Item(li_Row, Col_LastProcess) = ds.Tables(0).Rows(i)("LastProcess").ToString.Trim
                        grid.Item(li_Row, Col_NextProcess) = ds.Tables(0).Rows(i)("NextProcess").ToString.Trim

                        grid.GetCellRange(li_Row, Col_ProcessStatus).StyleNew.TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.CenterCenter

                        li_Row = li_Row + 1

                    Next

                End If

            End Using

        Catch ex As Exception
            WriteToErrorLog("Load data grid error", ex.Message)
        End Try
    End Sub

    Public Sub up_FillBackColourGrid(ByVal grid As C1FlexGrid, ByVal toCol As Integer, ByVal pRow As Integer)
        Dim csHeader As C1.Win.C1FlexGrid.CellStyle
        Dim csDetail As C1.Win.C1FlexGrid.CellStyle

        If pRow = 0 Then
            csHeader = grid.Styles.Add("BackColorHeader")
            csHeader.BackColor = Drawing.Color.Gainsboro
            csHeader.ForeColor = Drawing.Color.Black
            Dim rg As CellRange = grid.GetCellRange(0, 0, 0, toCol)
            rg.Style = csHeader
        Else
            csDetail = grid.Styles.Add("BackColorDetail")
            csDetail.BackColor = Drawing.Color.White
            Dim rgDetail As CellRange = grid.GetCellRange(pRow, toCol)
            rgDetail.Style = csDetail
        End If

    End Sub

    Private Sub up_Process()
        Dim pProses As String = ""
        Try

            ''Copy File From FTP
            ''==================
            pProses = "Copy and Download file from FTP"
            up_FTP()
            ''==================

            '01. PACKING M01 - SHRINK CUTTING
            pProses = "PACKING M01 - SHRINK CUTTING"
            up_ProcessData(ShrinkCutting1, "PACKING M01 - SHRINK CUTTING", gs_LocalPathShrinkCutting1, gs_FileName_QualityShrinkCutting1, 0, "", "")

            '02. PACKING M01 - WEIGHT BATCH
            pProses = "PACKING M01 - WEIGHT BATCH"
            up_ProcessData(WeightBatch1, "PACKING M01 - WEIGHT BATCH", gs_LocalPathWeightBatch1, gs_FileName_QualityWeightBatch1, 0, "", "")

            '03. PACKING M01 - HEATER
            pProses = "PACKING M01 - HEATER"
            up_ProcessData(Heater1, "PACKING M01 - HEATER", gs_LocalPathHeater1, gs_FileName_QualityHeater1, 0, "", "")

            '04. PACKING M01 - TAPING 3M
            pProses = "PACKING M01 - TAPING 3M"
            up_ProcessData(Taping1, "V", gs_LocalPathTaping1, gs_FileName_QualityTaping1, 0, "", "")

            '05. PACKING M02 - WEIGHT BOX
            pProses = "PACKING M02 - WEIGHT BOX"
            up_ProcessData(WeightBox2, "PACKING M02 - WEIGHT BOX", gs_LocalPathWeightBox2, gs_FileName_QualityWeightBox2, 0, "", "")

            '02. PACKING M02 - SHRINK CUTTING
            pProses = "PACKING M02 - SHRINK CUTTING"
            up_ProcessData(ShrinkCutting2, "PACKING M02 - SHRINK CUTTING", gs_LocalPathShrinkCutting2, gs_FileName_QualityShrinkCutting2, 0, "", "")

            '02. PACKING M02 - WEIGHT BATCH
            pProses = "PACKING M02 - WEIGHT BATCH"
            up_ProcessData(WeightBatch2, "PACKING M02 - WEIGHT BATCH", gs_LocalPathWeightBatch2, gs_FileName_QualityWeightBatch2, 0, "", "")

            '03. PACKING M02 - HEATER
            pProses = "PACKING M02 - HEATER"
            up_ProcessData(Heater2, "PACKING M02 - HEATER", gs_LocalPathHeater2, gs_FileName_QualityHeater2, 0, "", "")

            '04. PACKING M02 - TAPING 3M
            pProses = "PACKING M02 - TAPING 3M"
            up_ProcessData(Taping2, "V", gs_LocalPathTaping2, gs_FileName_QualityTaping2, 0, "", "")

            '05. PACKING M02 - WEIGHT BOX
            pProses = "PACKING M02 - WEIGHT BOX"
            up_ProcessData(WeightBox2, "PACKING M02 - WEIGHT BOX", gs_LocalPathWeightBox2, gs_FileName_QualityWeightBox2, 0, "", "")

            '11. PILLOW PROCESS
            pProses = "PILLOW PROCESS"
            up_ProcessData(PillowProcess, "PILLOW PROCESS", gs_LocalPathPillowProcess, gs_FileName_QualityPillowProcess, 0, "", "")

        Catch ex As Exception
            WriteToErrorLog(pProses & "error : ", ex.Message)
        Finally
            txtMsg.Text = "Data OK"
        End Try

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
                cmd.Parameters.AddWithValue("Type", "ALL")

                da = New SqlDataAdapter(cmd)
                da.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then

                    up_ClearVariable()

                    For i As Integer = 0 To ds.Tables(0).Rows.Count - 1

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "PACKING M01 - SHRINK CUTTING" Then
                            gs_ServerPathShrinkCutting1 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathShrinkCutting1 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserShrinkCutting1 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordShrinkCutting1 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalShrinkCutting1 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "PACKING M01 - WEIGHT BATCH" Then
                            gs_ServerPathWeightBatch1 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathWeightBatch1 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserWeightBatch1 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordWeightBatch1 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalWeightBatch1 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "PACKING M01 - HEATER" Then
                            gs_ServerPathHeater1 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathHeater1 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserHeater1 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordHeater1 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalHeater1 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "PACKING M01 - TAPING 3M" Then
                            gs_ServerPathTaping1 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathTaping1 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserTaping1 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordTaping1 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalTaping1 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "PACKING M01 - WEIGHT BOX" Then
                            gs_ServerPathWeightBox1 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathWeightBox1 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserWeightBox1 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordWeightBox1 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalWeightBox1 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "PACKING M02 - SHRINK CUTTING" Then
                            gs_ServerPathShrinkCutting2 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathShrinkCutting2 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserShrinkCutting2 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordShrinkCutting2 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalShrinkCutting2 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "PACKING M02 - WEIGHT BATCH" Then
                            gs_ServerPathWeightBatch2 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathWeightBatch2 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserWeightBatch2 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordWeightBatch2 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalWeightBatch2 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "PACKING M02 - HEATER" Then
                            gs_ServerPathHeater2 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathHeater2 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserHeater2 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordHeater2 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalHeater2 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "PACKING M02 - TAPING 3M" Then
                            gs_ServerPathTaping2 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathTaping2 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserTaping2 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordTaping2 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalTaping2 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "PACKING M02 - WEIGHT BOX" Then
                            gs_ServerPathWeightBox2 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathWeightBox2 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserWeightBox2 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordWeightBox2 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalWeightBox2 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "PILLOW PROCESS" Then
                            gs_ServerPathPillowProcess = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathPillowProcess = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserPillowProcess = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordPillowProcess = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalPillowProcess = ds.Tables(0).Rows(i)("TimeInterval")
                        End If
                    Next

                End If

            End Using

        Catch ex As Exception
            WriteToErrorLog("Load data FTP error", ex.Message)
        End Try

    End Sub

    Private Sub up_ClearVariable()

        gs_ServerPathShrinkCutting1 = ""
        gs_LocalPathShrinkCutting1 = ""
        gs_UserShrinkCutting1 = ""
        gs_PasswordShrinkCutting1 = ""

        gs_ServerPathWeightBatch1 = ""
        gs_LocalPathWeightBatch1 = ""
        gs_UserWeightBatch1 = ""
        gs_PasswordWeightBatch1 = ""

        gs_ServerPathHeater1 = ""
        gs_LocalPathHeater1 = ""
        gs_UserHeater1 = ""
        gs_PasswordHeater1 = ""

        gs_ServerPathTaping1 = ""
        gs_LocalPathTaping1 = ""
        gs_UserTaping1 = ""
        gs_PasswordTaping1 = ""

        gs_ServerPathWeightBox1 = ""
        gs_LocalPathWeightBox1 = ""
        gs_UserWeightBox1 = ""
        gs_PasswordWeightBox1 = ""

        gs_ServerPathShrinkCutting2 = ""
        gs_LocalPathShrinkCutting2 = ""
        gs_UserShrinkCutting2 = ""
        gs_PasswordShrinkCutting2 = ""

        gs_ServerPathWeightBatch2 = ""
        gs_LocalPathWeightBatch2 = ""
        gs_UserWeightBatch2 = ""
        gs_PasswordWeightBatch2 = ""

        gs_ServerPathHeater2 = ""
        gs_LocalPathHeater2 = ""
        gs_UserHeater2 = ""
        gs_PasswordHeater2 = ""

        gs_ServerPathTaping2 = ""
        gs_LocalPathTaping2 = ""
        gs_UserTaping2 = ""
        gs_PasswordTaping2 = ""

        gs_ServerPathWeightBox2 = ""
        gs_LocalPathWeightBox2 = ""
        gs_UserWeightBox2 = ""
        gs_PasswordWeightBox2 = ""

        gs_ServerPathPillowProcess = ""
        gs_LocalPathPillowProcess = ""
        gs_UserPillowProcess = ""
        gs_PasswordPillowProcess = ""
    End Sub

#Region "PROCESS"

    Private Sub up_TimeStart()
        m_Finish = False
        Me.Cursor = Cursors.WaitCursor

        Try

            Thread.Sleep(200)
            Thd_ShrinkCutting1 = New SchedulerSetting
            With Thd_ShrinkCutting1
                .Name = "ShrinkCutting1"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalShrinkCutting1
                .ScheduleThd = New Thread(AddressOf up_RefreshShrinkCutting1)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "ShrinkCutting1"
                .ScheduleThd.Start()
            End With

            Thread.Sleep(220)
            Thd_WeightBatch1 = New SchedulerSetting
            With Thd_WeightBatch1
                .Name = "WeightBatch1"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalWeightBatch1
                .ScheduleThd = New Thread(AddressOf up_RefreshWeightBatch1)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "WeightBatch1"
                .ScheduleThd.Start()
            End With

            Thread.Sleep(240)
            Thd_Heater1 = New SchedulerSetting
            With Thd_Heater1
                .Name = "Heater1"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalHeater1
                .ScheduleThd = New Thread(AddressOf up_RefreshHeater1)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "Heater1"
                .ScheduleThd.Start()
            End With

            Thread.Sleep(260)
            Thd_Taping1 = New SchedulerSetting
            With Thd_Taping1
                .Name = "Taping1"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalTaping1
                .ScheduleThd = New Thread(AddressOf up_RefreshTaping1)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "Taping1"
                .ScheduleThd.Start()
            End With

            Thread.Sleep(280)
            Thd_WeightBox1 = New SchedulerSetting
            With Thd_WeightBox1
                .Name = "WeightBox1"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalWeightBox1
                .ScheduleThd = New Thread(AddressOf up_RefreshWeightBox1)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "WeightBox1"
                .ScheduleThd.Start()
            End With

            Thread.Sleep(200)
            Thd_ShrinkCutting2 = New SchedulerSetting
            With Thd_ShrinkCutting2
                .Name = "ShrinkCutting2"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalShrinkCutting2
                .ScheduleThd = New Thread(AddressOf up_RefreshShrinkCutting2)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "ShrinkCutting2"
                .ScheduleThd.Start()
            End With

            Thread.Sleep(220)
            Thd_WeightBatch2 = New SchedulerSetting
            With Thd_WeightBatch2
                .Name = "WeightBatch2"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalWeightBatch2
                .ScheduleThd = New Thread(AddressOf up_RefreshWeightBatch2)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "WeightBatch2"
                .ScheduleThd.Start()
            End With

            Thread.Sleep(240)
            Thd_Heater2 = New SchedulerSetting
            With Thd_Heater2
                .Name = "Heater2"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalHeater2
                .ScheduleThd = New Thread(AddressOf up_RefreshHeater2)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "Heater2"
                .ScheduleThd.Start()
            End With

            Thread.Sleep(260)
            Thd_Taping2 = New SchedulerSetting
            With Thd_Taping2
                .Name = "Taping2"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalTaping2
                .ScheduleThd = New Thread(AddressOf up_RefreshTaping2)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "Taping2"
                .ScheduleThd.Start()
            End With

            Thread.Sleep(280)
            Thd_WeightBox2 = New SchedulerSetting
            With Thd_WeightBox2
                .Name = "WeightBox2"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalWeightBox2
                .ScheduleThd = New Thread(AddressOf up_RefreshWeightBox2)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "WeightBox2"
                .ScheduleThd.Start()
            End With

            Thread.Sleep(280)
            Thd_PillowProcess = New SchedulerSetting
            With Thd_PillowProcess
                .Name = "PillowProcess"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalPillowProcess
                .ScheduleThd = New Thread(AddressOf up_RefreshPillowProcess)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "PillowProcess"
                .ScheduleThd.Start()
            End With
        Catch ex As Exception
            txtMsg.Text = ex.Message
            WriteToErrorLog("TimeStart", ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub up_TimeStop()

        m_Finish = True

        SyncLock Thd_ShrinkCutting1.Lock
            Thd_ShrinkCutting1.EndSchedule = True
        End SyncLock

        SyncLock Thd_WeightBatch1.Lock
            Thd_WeightBatch1.EndSchedule = True
        End SyncLock

        SyncLock Thd_Heater1.Lock
            Thd_Heater1.EndSchedule = True
        End SyncLock

        SyncLock Thd_Taping1.Lock
            Thd_Taping1.EndSchedule = True
        End SyncLock

        SyncLock Thd_WeightBox1.Lock
            Thd_WeightBox1.EndSchedule = True
        End SyncLock

        SyncLock Thd_ShrinkCutting2.Lock
            Thd_ShrinkCutting2.EndSchedule = True
        End SyncLock

        SyncLock Thd_WeightBatch2.Lock
            Thd_WeightBatch2.EndSchedule = True
        End SyncLock

        SyncLock Thd_Heater2.Lock
            Thd_Heater2.EndSchedule = True
        End SyncLock

        SyncLock Thd_Taping2.Lock
            Thd_Taping2.EndSchedule = True
        End SyncLock

        SyncLock Thd_WeightBox2.Lock
            Thd_WeightBox2.EndSchedule = True
        End SyncLock

        SyncLock Thd_PillowProcess.Lock
            Thd_PillowProcess.EndSchedule = True
        End SyncLock
    End Sub

    Private Function up_ToInserDatatable_Trouble(ByVal pLocalPath As String, ByVal pFileName As String, ByVal pLineCode As String) As DataTable
        Dim dt As New DataTable
        Dim tmpDatde As String()
        Dim Col_Line As String = "", AlarmCode As String
        Dim pAlarm1 As String = "", EndTime As String, LastUpdate As String
        Dim pMachine As String = "", StartTime As String = "", ModeCls As String = "", StatusCls As String = ""
        Dim li_Add As Integer = 0

        Try

            Dim di As New IO.DirectoryInfo(pLocalPath)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
            Dim fi As IO.FileInfo = Nothing

            Dim x As Integer
            Dim strarray(1, 1) As String
            Dim pCls As String = ""

            For Each fi In aryFi

                If fi.Name = pFileName Then

                    Dim dtCSV As DataTable = ReadCSV(fi.FullName)

                    If dtCSV.Rows.Count > 0 Then

                        With dt.Columns
                            .Add("LineCode", GetType(String))
                            .Add("Mode", GetType(String))
                            .Add("Status", GetType(String))
                            .Add("TroubleCode", GetType(String))
                            .Add("StartTime", GetType(String))
                            .Add("EndTime", GetType(String))
                            .Add("LstUpdate", GetType(String))
                            .Add("LastUser", GetType(String))
                        End With

                        For x = 0 To dtCSV.Rows.Count - 2
                            tmpDatde = Split(Trim(dtCSV.Rows(x)(0)), "/")
                            Col_Line = pLineCode
                            If tmpDatde(0) > 0 Then

                                StartTime = Format(CDate(20 & Trim(dtCSV.Rows(x)(0)) & " " & Trim(dtCSV.Rows(x)(1))), "yyyy-MM-dd HH:mm:ss")
                                ModeCls = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(2))), 0, 15)
                                StatusCls = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(3))), 0, 15)
                                LastUpdate = Format(Now, "yyyy-MM-dd HH:mm:ss")

                                If StatusCls = 0 Then

                                    pAlarm1 = 0

                                    If Trim(dtCSV.Rows(x + 1)(0)) <> "" Then
                                        EndTime = Format(CDate(20 & Trim(dtCSV.Rows(x + 1)(0)) & " " & Trim(dtCSV.Rows(x + 1)(1))), "yyyy-MM-dd HH:mm:ss")
                                    Else
                                        EndTime = Format(CDate(20 & Trim(dtCSV.Rows(x)(0)) & " " & Trim(dtCSV.Rows(x)(1))), "yyyy-MM-dd HH:mm:ss")
                                    End If

                                    dt.Rows.Add(Col_Line, ModeCls, StatusCls, pAlarm1, StartTime, EndTime, LastUpdate, "Tes")

                                Else

                                    If x = dtCSV.Rows.Count - 2 Then
                                        li_Add = 0
                                    Else
                                        li_Add = 1
                                    End If

                                    pAlarm1 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(4))), 0, 15)
                                    EndTime = uf_ProcessAlarm_Trouble(dtCSV, x + li_Add, pAlarm1)

                                    dt.Rows.Add(Col_Line, ModeCls, StatusCls, pAlarm1, StartTime, EndTime, LastUpdate, "Tes")

                                End If

                            End If
                        Next

                        'Send to History Folder
                        '======================
                        My.Computer.FileSystem.MoveFile(pLocalPath & "\" & fi.Name, pLocalPath & "\History\" & fi.Name, True)

                        Return dt

                    End If

                End If

            Next

        Catch ex As Exception
            WriteToErrorLog("Trouble Process error in LineCode : " & pLineCode, ex.Message)
        End Try

    End Function

    Private Function up_ToInserDatatable_Quality(ByVal pLocalPath As String, ByVal pFileName As String, ByVal pLineCode As String, ByVal pMCCode As String) As DataTable
        Dim dt As New DataTable

        Dim col_Line, col_Machine, col_model, Col_MeasTime As String
        Dim tmpDatde As String()
        Dim data_001, data_002, data_003, data_004, data_005 As String
        Dim data_006, data_007, data_008, data_009, data_010 As String
        Dim data_011, data_012, data_013, data_014, data_015 As String
        Dim data_016, data_017, data_018, data_019, data_020 As String
        Dim data_021, data_022, data_023, data_024, data_025 As String

        Try

            Dim di As New IO.DirectoryInfo(pLocalPath)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
            Dim fi As IO.FileInfo = Nothing

            Dim x As Integer
            Dim strarray(1, 1) As String
            Dim pCls As String = ""

            For Each fi In aryFi

                If fi.Name = pFileName Then

                    Dim dtCSV As DataTable = ReadCSV(fi.FullName)

                    'Backup CSV on database
                    up_SavedataCSV(dtCSV, pLineCode)

                    If dtCSV.Rows.Count > 0 Then

                        With dt.Columns
                            .Add("LineCode", GetType(String))
                            .Add("Machine", GetType(String))
                            .Add("Model", GetType(String))
                            .Add("MeasTime", GetType(String))
                            .Add("data_001", GetType(String))
                            .Add("data_002", GetType(String))
                            .Add("data_003", GetType(String))
                            .Add("data_004", GetType(String))
                            .Add("data_005", GetType(String))
                            .Add("data_006", GetType(String))
                            .Add("data_007", GetType(String))
                            .Add("data_008", GetType(String))
                            .Add("data_009", GetType(String))
                            .Add("data_010", GetType(String))
                            .Add("data_011", GetType(String))
                            .Add("data_012", GetType(String))
                            .Add("data_013", GetType(String))
                            .Add("data_014", GetType(String))
                            .Add("data_015", GetType(String))
                            .Add("data_016", GetType(String))
                            .Add("data_017", GetType(String))
                            .Add("data_018", GetType(String))
                            .Add("data_019", GetType(String))
                            .Add("data_020", GetType(String))
                            .Add("data_021", GetType(String))
                            .Add("data_022", GetType(String))
                            .Add("data_023", GetType(String))
                            .Add("data_024", GetType(String))
                            .Add("data_025", GetType(String))
                        End With

                        For x = 0 To dtCSV.Rows.Count - 2
                            tmpDatde = Split(Trim(dtCSV.Rows(x)(0)), "/")
                            If tmpDatde(0) > 0 Then
                                col_Line = pLineCode
                                col_Machine = pMCCode
                                col_model = "BBRSRUSA0PAD"
                                Col_MeasTime = Format(CDate(20 & Trim(dtCSV.Rows(x)(0)) & " " & Trim(dtCSV.Rows(x)(1))), "yyyy-MM-dd HH:mm:ss")
                                data_001 = "2"
                                data_002 = "0"
                                data_003 = "0"
                                data_004 = "0"
                                data_005 = "0"
                                data_006 = "0"
                                data_007 = "0"
                                data_008 = "0"
                                data_009 = "0"
                                data_010 = "0"
                                data_011 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(2))), 0, 15)
                                data_012 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(3))), 0, 15)
                                data_013 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(4))), 0, 15)
                                data_014 = Trim(dtCSV.Rows(x)(5))
                                data_015 = Trim(dtCSV.Rows(x)(6))
                                data_016 = Trim(dtCSV.Rows(x)(7))
                                data_017 = Trim(dtCSV.Rows(x)(8))
                                data_018 = Trim(dtCSV.Rows(x)(9))
                                data_019 = Trim(dtCSV.Rows(x)(10))
                                data_020 = Trim(dtCSV.Rows(x)(11))
                                data_021 = Trim(dtCSV.Rows(x)(12))
                                data_022 = Trim(dtCSV.Rows(x)(13))
                                data_023 = Trim(dtCSV.Rows(x)(14))
                                data_024 = Trim(dtCSV.Rows(x)(15))
                                data_025 = Trim(dtCSV.Rows(x)(16))

                                dt.Rows.Add(col_Line, col_Machine, col_model, Col_MeasTime,
                                            data_001, data_002, data_003, data_004, data_005,
                                            data_006, data_007, data_008, data_009, data_010,
                                            data_011, data_012, data_013, data_014, data_015,
                                            data_016, data_017, data_018, data_019, data_020,
                                            data_021, data_022, data_023, data_024, data_025)

                            End If
                        Next

                        'Send to History Folder
                        '======================
                        My.Computer.FileSystem.MoveFile(pLocalPath & "\" & fi.Name, pLocalPath & "\History\" & fi.Name, True)

                        Return dt


                    End If

                End If

            Next

            Return dt

        Catch ex As Exception
            txtMsg.Text = ex.Message
            WriteToErrorLog("Quality Process error in LineCode : " & pLineCode, ex.Message)
        End Try

    End Function

    Private Sub up_RefreshShrinkCutting1()
        up_Refresh(ShrinkCutting1, "PACKING M01 - Shrink Cutting", Thd_ShrinkCutting1, gs_ServerPathShrinkCutting1, gs_LocalPathShrinkCutting1, gs_UserShrinkCutting1, gs_PasswordShrinkCutting1, gs_FileName_QualityShrinkCutting1, 0, "", "")
    End Sub
    Private Sub up_RefreshWeightBatch1()
        up_Refresh(WeightBatch1, "PACKING M01 - Weight Batch", Thd_WeightBatch1, gs_ServerPathWeightBatch1, gs_LocalPathWeightBatch1, gs_UserWeightBatch1, gs_PasswordWeightBatch1, gs_FileName_QualityWeightBatch1, 0, "", "")
    End Sub
    Private Sub up_RefreshHeater1()
        up_Refresh(Heater1, "PACKING M01 - Heater", Thd_Heater1, gs_ServerPathHeater1, gs_LocalPathHeater1, gs_UserHeater1, gs_PasswordHeater1, gs_FileName_QualityHeater1, 0, "", "")
    End Sub
    Private Sub up_RefreshTaping1()
        up_Refresh(Taping1, "PACKING M01 - Taping 3M", Thd_Taping1, gs_ServerPathTaping1, gs_LocalPathTaping1, gs_UserTaping1, gs_PasswordTaping1, gs_FileName_QualityTaping1, 0, "", "")
    End Sub
    Private Sub up_RefreshWeightBox1()
        up_Refresh(WeightBox1, "PACKING M01 - Weight Box", Thd_WeightBox1, gs_ServerPathWeightBox1, gs_LocalPathWeightBox1, gs_UserWeightBox1, gs_PasswordWeightBox1, gs_FileName_QualityWeightBox1, 0, "", "")
    End Sub
    Private Sub up_RefreshShrinkCutting2()
        up_Refresh(ShrinkCutting2, "PACKING M02 - Shrink Cutting", Thd_ShrinkCutting2, gs_ServerPathShrinkCutting2, gs_LocalPathShrinkCutting2, gs_UserShrinkCutting2, gs_PasswordShrinkCutting2, gs_FileName_QualityShrinkCutting2, 0, "", "")
    End Sub
    Private Sub up_RefreshWeightBatch2()
        up_Refresh(WeightBatch2, "PACKING M02 - Weight Batch", Thd_WeightBatch2, gs_ServerPathWeightBatch2, gs_LocalPathWeightBatch2, gs_UserWeightBatch2, gs_PasswordWeightBatch2, gs_FileName_QualityWeightBatch2, 0, "", "")
    End Sub
    Private Sub up_RefreshHeater2()
        up_Refresh(Heater2, "PACKING M02 - Heater", Thd_Heater2, gs_ServerPathHeater2, gs_LocalPathHeater2, gs_UserHeater2, gs_PasswordHeater2, gs_FileName_QualityHeater2, 0, "", "")
    End Sub
    Private Sub up_RefreshTaping2()
        up_Refresh(Taping2, "PACKING M02 - Taping 3M", Thd_Taping2, gs_ServerPathTaping2, gs_LocalPathTaping2, gs_UserTaping2, gs_PasswordTaping2, gs_FileName_QualityTaping2, 0, "", "")
    End Sub
    Private Sub up_RefreshWeightBox2()
        up_Refresh(WeightBox2, "PACKING M02 - Weight Box", Thd_WeightBox2, gs_ServerPathWeightBox2, gs_LocalPathWeightBox2, gs_UserWeightBox2, gs_PasswordWeightBox2, gs_FileName_QualityWeightBox2, 0, "", "")
    End Sub
    Private Sub up_RefreshPillowProcess()
        up_Refresh(PillowProcess, "Pillow Process", Thd_PillowProcess, gs_ServerPathPillowProcess, gs_LocalPathPillowProcess, gs_UserPillowProcess, gs_PasswordPillowProcess, gs_FileName_QualityPillowProcess, 0, "", "")
    End Sub
    Private Sub up_Refresh(ByVal pProcess As Integer, ByVal pProcessName As String, ByVal Thd As SchedulerSetting, ByVal pServerPath As String, ByVal pLocalPath As String, ByVal pUser As String, ByVal pPassword As String, ByVal pFileNameHistory As String, ByVal pGroupCount As Integer, ByVal pLineCode As String, ByVal pMCCode As String)
        Dim errMsg As String = ""
        Dim i_Data As String = ""
        Dim DelayTime As Integer = 0
        Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
        Dim startime As DateTime

        Do Until m_Finish
            Try

                Application.DoEvents()
                grid.Item(pProcess, Col_ProcessStatus) = "RUNNING"

                Thd.Status = "In Progress"
                startime = Now

                'FTP Weight Check
                '====================
                If pServerPath <> "" And pLocalPath <> "" Then
                    FilesList = GetFtpFileList(pServerPath, pUser, pPassword)
                    If FilesList.Count > 0 Then
                        'Download File
                        '=============
                        DownloadFiles(FilesList, "", pUser, pPassword, pServerPath, pLocalPath, pFileNameHistory, "zz")

                        'Delete File
                        '===========
                        DeleteFiles(FilesList, "", pUser, pPassword, pServerPath, pLocalPath, pFileNameHistory, "zz")
                    End If
                End If
                '====================

                up_ProcessData(pProcess, pProcessName, pLocalPath, pFileNameHistory, pGroupCount, pLineCode, pMCCode)

                Threading.Thread.Sleep(DelayTime)

            Catch ex As Exception
                grid.Item(pProcess, Col_ErrorMessage) = ex.Message
                WriteToErrorLog($"{pProcessName} Scheduler", ex.Message)
            Finally
                Application.DoEvents()
                grid.Item(pProcess, Col_ProcessStatus) = "IDDLE"
                Thd.Status = "iddle"
                Thread.Sleep(Thd.DelayTime)
            End Try

            SyncLock Thd.Lock
                If Thd.EndSchedule Then
                    m_Finish = True
                End If
            End SyncLock

        Loop
    End Sub

    Private Sub up_ProcessData(ByVal pProcess As Integer, ByVal pProcessName As String, ByVal pLocalPath As String, ByVal pFileNameHistory As String, ByVal pGroupCount As Integer, ByVal pLineCode As String, ByVal pMCCode As String)
        Dim con As New SqlConnection
        'gs_LocalPathMIX = "D:\PECGI CSV\Tamping\LOG_Mc21001"

        grid.Item(pProcess, Col_ErrorMessage) = ""

        Dim dtHis, dtInfo As New DataTable
        'dtHis = up_ToInserDatatable_Trouble2(pLocalPath, pFileNameHistory, "057")
        dtInfo = up_ToInserDatatable_Quality(pLocalPath, pFileNameHistory, pLineCode, pMCCode)

        con = New SqlConnection(ConStr)
        con.Open()

        Dim cmd As SqlCommand
        Dim SQLTrans As SqlTransaction

        SQLTrans = con.BeginTransaction

        Try

            Application.DoEvents()
            grid.Item(pProcess, Col_ProcessStatus) = "RUNNING"

            If dtHis IsNot Nothing Then
                If dtHis.Rows.Count > 0 Then

                    'clsSchedulerCSV_BA_DB.InsertData_History(dtMixHis, "MIX_His_CSV_")

                    cmd = New SqlCommand("sp_Insert_Log_CSV_Trouble_New", con)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Connection = con
                    cmd.Transaction = SQLTrans

                    Dim paramTbl As New SqlParameter()
                    paramTbl.ParameterName = "@LogDataTrouble"
                    paramTbl.SqlDbType = SqlDbType.Structured

                    Dim paramName As New SqlParameter()
                    paramName.ParameterName = "@ProcessTime"
                    paramName.SqlDbType = SqlDbType.DateTime

                    cmd.Parameters.Add(paramTbl)
                    cmd.Parameters.Add(paramName)

                    cmd.Parameters("@LogDataTrouble").Value = dtHis
                    cmd.Parameters("@ProcessTime").Value = Format(Now, "yyyy-MM-dd HH:mm:ss")

                    cmd.CommandTimeout = 100000
                    Dim i As Integer = cmd.ExecuteNonQuery()

                End If
            End If

            If dtInfo IsNot Nothing Then
                If dtInfo.Rows.Count > 0 Then

                    'clsSchedulerCSV_BA_DB.InsertData_Info(dtMixInfo, "MIX_Info_CSV_")
                    cmd = New SqlCommand("sp_Insert_Meas_CSV_Quality_New", con)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Connection = con
                    cmd.Transaction = SQLTrans

                    Dim paramTbl As New SqlParameter()
                    paramTbl.ParameterName = "@MeasData"
                    paramTbl.SqlDbType = SqlDbType.Structured

                    Dim paramName As New SqlParameter()
                    paramName.ParameterName = "@ProcessTime"
                    paramName.SqlDbType = SqlDbType.DateTime

                    cmd.Parameters.Add(paramTbl)
                    cmd.Parameters.Add(paramName)

                    cmd.Parameters("@MeasData").Value = dtInfo
                    cmd.Parameters("@ProcessTime").Value = Format(Now, "yyyy-MM-dd HH:mm:ss")

                    cmd.CommandTimeout = 100000
                    Dim i As Integer = cmd.ExecuteNonQuery()
                End If
            End If

            SQLTrans.Commit()

            Application.DoEvents()
            grid.Item(pProcess, Col_LastProcess) = Format(Now, "dd MMM yyyy HH:mm:ss")
            grid.Item(pProcess, Col_NextProcess) = Format(DateAdd(DateInterval.Minute, 5, Now), "dd MMM yyyy HH:mm:ss")

        Catch ex As Exception
            grid.Item(pProcess, Col_ErrorMessage) = ex.Message
            WriteToErrorLog($"{pProcessName} Process", ex.Message)
            SQLTrans.Rollback()
        End Try

    End Sub

    Private Sub up_FTP()

        'FTP Packing M01 - Shrink Cutting
        '====================
        If gs_ServerPathShrinkCutting1 <> "" And gs_LocalPathShrinkCutting1 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathShrinkCutting1, gs_UserShrinkCutting1, gs_PasswordShrinkCutting1)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserShrinkCutting1, gs_PasswordShrinkCutting1, gs_ServerPathShrinkCutting1, gs_LocalPathShrinkCutting1, gs_FileName_QualityShrinkCutting1, "zz")
            End If
        End If
        '====================

        'FTP Packing M01 - Weight Batch
        '====================
        If gs_ServerPathWeightBatch1 <> "" And gs_LocalPathWeightBatch1 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathWeightBatch1, gs_UserWeightBatch1, gs_PasswordWeightBatch1)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserWeightBatch1, gs_PasswordWeightBatch1, gs_ServerPathWeightBatch1, gs_LocalPathWeightBatch1, gs_FileName_QualityWeightBatch1, "zz")
            End If
        End If
        '====================

        'FTP Packing M01 - Heater
        '====================
        If gs_ServerPathHeater1 <> "" And gs_LocalPathHeater1 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathHeater1, gs_UserHeater1, gs_PasswordHeater1)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserHeater1, gs_PasswordHeater1, gs_ServerPathHeater1, gs_LocalPathHeater1, gs_FileName_QualityHeater1, "zz")
            End If
        End If
        '====================

        'FTP Packing M01 - Taping 3M
        '====================
        If gs_ServerPathTaping1 <> "" And gs_LocalPathTaping1 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathTaping1, gs_UserTaping1, gs_PasswordTaping1)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserTaping1, gs_PasswordTaping1, gs_ServerPathTaping1, gs_LocalPathTaping1, gs_FileName_QualityTaping1, "zz")
            End If
        End If
        '====================

        'FTP Packing M01 - Weight Box
        '====================
        If gs_ServerPathWeightBox1 <> "" And gs_LocalPathWeightBox1 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathWeightBox1, gs_UserWeightBox1, gs_PasswordWeightBox1)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserWeightBox1, gs_PasswordWeightBox1, gs_ServerPathWeightBox1, gs_LocalPathWeightBox1, gs_FileName_QualityWeightBox1, "zz")
            End If
        End If
        '====================

        'FTP Packing M02 - Shrink Cutting
        '====================
        If gs_ServerPathShrinkCutting2 <> "" And gs_LocalPathShrinkCutting2 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathShrinkCutting2, gs_UserShrinkCutting2, gs_PasswordShrinkCutting2)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserShrinkCutting2, gs_PasswordShrinkCutting2, gs_ServerPathShrinkCutting2, gs_LocalPathShrinkCutting2, gs_FileName_QualityShrinkCutting2, "zz")
            End If
        End If
        '====================

        'FTP Packing M02 - Weight Batch
        '====================
        If gs_ServerPathWeightBatch2 <> "" And gs_LocalPathWeightBatch2 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathWeightBatch2, gs_UserWeightBatch2, gs_PasswordWeightBatch2)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserWeightBatch2, gs_PasswordWeightBatch2, gs_ServerPathWeightBatch2, gs_LocalPathWeightBatch2, gs_FileName_QualityWeightBatch2, "zz")
            End If
        End If
        '====================

        'FTP Packing M02 - Heater
        '====================
        If gs_ServerPathHeater2 <> "" And gs_LocalPathHeater2 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathHeater2, gs_UserHeater2, gs_PasswordHeater2)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserHeater2, gs_PasswordHeater2, gs_ServerPathHeater2, gs_LocalPathHeater2, gs_FileName_QualityHeater2, "zz")
            End If
        End If
        '====================

        'FTP Packing M02 - Taping 3M
        '====================
        If gs_ServerPathTaping2 <> "" And gs_LocalPathTaping2 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathTaping2, gs_UserTaping2, gs_PasswordTaping2)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserTaping2, gs_PasswordTaping2, gs_ServerPathTaping2, gs_LocalPathTaping2, gs_FileName_QualityTaping2, "zz")
            End If
        End If
        '====================

        'FTP Packing M02 - Weight Box
        '====================
        If gs_ServerPathWeightBox2 <> "" And gs_LocalPathWeightBox2 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathWeightBox2, gs_UserWeightBox2, gs_PasswordWeightBox2)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserWeightBox2, gs_PasswordWeightBox2, gs_ServerPathWeightBox2, gs_LocalPathWeightBox2, gs_FileName_QualityWeightBox2, "zz")
            End If
        End If
        '====================

        'FTP Packing M02 - Pillow Process
        '====================
        If gs_ServerPathPillowProcess <> "" And gs_LocalPathPillowProcess <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathPillowProcess, gs_UserPillowProcess, gs_PasswordPillowProcess)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserPillowProcess, gs_PasswordPillowProcess, gs_ServerPathPillowProcess, gs_LocalPathPillowProcess, gs_FileName_QualityPillowProcess, "zz")
            End If
        End If
        '====================
    End Sub

    Private Sub up_SavedataCSV(ByVal dtTmp As DataTable, ByVal pLineCode As String)

        Dim con As New SqlConnection
        Dim dt As New DataTable

        With dt.Columns
            .Add("Alarm_Code_A", GetType(String))
            .Add("Alarm_Code_B", GetType(String))
            .Add("Alarm_Code_C", GetType(String))
            .Add("Alarm_Code_D", GetType(String))
            .Add("Alarm_Code_E", GetType(String))
            .Add("Alarm_Code_F", GetType(String))
            .Add("Alarm_Code_G", GetType(String))
            .Add("Alarm_Code_H", GetType(String))
            .Add("Alarm_Code_I", GetType(String))
            .Add("Alarm_Code_J", GetType(String))
            .Add("Alarm_Code_K", GetType(String))
            .Add("Alarm_Code_L", GetType(String))
            .Add("Alarm_Code_M", GetType(String))
            .Add("Alarm_Code_N", GetType(String))
            .Add("Alarm_Code_O", GetType(String))
            .Add("Alarm_Code_P", GetType(String))
            .Add("Alarm_Code_Q", GetType(String))
            .Add("Alarm_Code_R", GetType(String))
            .Add("Alarm_Code_S", GetType(String))
            .Add("Alarm_Code_T", GetType(String))
            .Add("Alarm_Code_U", GetType(String))
        End With
        'Format(CDate(20 & Trim(dtCSV.Rows(x)(0)) & " " & Trim(dtCSV.Rows(x)(1))), "yyyy-MM-dd HH:mm:ss")
        For x = 0 To dtTmp.Rows.Count - 1
            If Trim(dtTmp.Rows(x).Item(0)) <> "" Then
                dt.Rows.Add(Format(CDate(20 & (Trim(dtTmp.Rows(x).Item(0)))), "yyyy-MM-dd"), Format(CDate(dtTmp.Rows(x).Item(1)), "HH:mm:ss"), Trim(dtTmp.Rows(x).Item(2)), Trim(dtTmp.Rows(x).Item(3)),
                            Trim(dtTmp.Rows(x).Item(4)), Trim(dtTmp.Rows(x).Item(5)), Trim(dtTmp.Rows(x).Item(6)), Trim(dtTmp.Rows(x).Item(7)), Trim(dtTmp.Rows(x).Item(8)), Trim(dtTmp.Rows(x).Item(9)),
                            Trim(dtTmp.Rows(x).Item(10)), Trim(dtTmp.Rows(x).Item(11)), Trim(dtTmp.Rows(x).Item(12)), Trim(dtTmp.Rows(x).Item(13)), Trim(dtTmp.Rows(x).Item(14)), Trim(dtTmp.Rows(x).Item(15)),
                            Trim(dtTmp.Rows(x).Item(16)), Trim(dtTmp.Rows(x).Item(17)), Trim(dtTmp.Rows(x).Item(18)), Trim(dtTmp.Rows(x).Item(19)), Trim(dtTmp.Rows(x).Item(20)))
            End If
        Next

        con = New SqlConnection(ConStr)
        con.Open()

        Dim cmd As SqlCommand
        Dim SQLTrans As SqlTransaction

        SQLTrans = con.BeginTransaction

        Try

            cmd = New SqlCommand("sp_PLC_CSV_INFO_Quality", con)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = con
            cmd.Transaction = SQLTrans

            Dim paramTbl As New SqlParameter()
            paramTbl.ParameterName = "@InfoData"
            paramTbl.SqlDbType = SqlDbType.Structured

            Dim paramName As New SqlParameter()
            paramName.ParameterName = "@LineCode"
            paramName.SqlDbType = SqlDbType.VarChar

            cmd.Parameters.Add(paramTbl)
            cmd.Parameters.Add(paramName)

            cmd.Parameters("@InfoData").Value = dt
            cmd.Parameters("@LineCode").Value = pLineCode

            cmd.CommandTimeout = 100000
            Dim i As Integer = cmd.ExecuteNonQuery()



            SQLTrans.Commit()

        Catch ex As Exception
            WriteToErrorLog("Kneading Process insert data CSV", ex.Message)
            SQLTrans.Rollback()
            txtMsg.Text = ex.Message
        End Try

    End Sub

    Public Function DecimalToBinary(DecimalNum As Long) As String
        Dim tmp As String
        Dim n As Long

        n = DecimalNum

        tmp = Trim(Str(n Mod 2))
        n = n \ 2

        Do While n <> 0
            tmp = Trim(Str(n Mod 2)) & tmp
            n = n \ 2
        Loop

        If Len(tmp) = 1 Then
            tmp = "000000000000000" & tmp
        ElseIf Len(tmp) = 2 Then
            tmp = "00000000000000" & tmp
        ElseIf Len(tmp) = 3 Then
            tmp = "0000000000000" & tmp
        ElseIf Len(tmp) = 4 Then
            tmp = "000000000000" & tmp
        ElseIf Len(tmp) = 5 Then
            tmp = "00000000000" & tmp
        ElseIf Len(tmp) = 6 Then
            tmp = "0000000000" & tmp
        ElseIf Len(tmp) = 7 Then
            tmp = "000000000" & tmp
        ElseIf Len(tmp) = 8 Then
            tmp = "00000000" & tmp
        ElseIf Len(tmp) = 9 Then
            tmp = "0000000" & tmp
        ElseIf Len(tmp) = 10 Then
            tmp = "000000" & tmp
        ElseIf Len(tmp) = 11 Then
            tmp = "00000" & tmp
        ElseIf Len(tmp) = 12 Then
            tmp = "0000" & tmp
        ElseIf Len(tmp) = 13 Then
            tmp = "000" & tmp
        ElseIf Len(tmp) = 14 Then
            tmp = "00" & tmp
        ElseIf Len(tmp) = 15 Then
            tmp = "0" & tmp
        End If

        DecimalToBinary = tmp

    End Function

    Public Function up_GetCodeTrouble(ByVal pTmp As String, ByVal pStart As Integer, ByVal pEnd As Integer) As String

        Dim pChr As String = ""
        Dim pMid As Integer = 16
        Dim pCode As String = "0"

        For i As Integer = pStart To pEnd
            pChr = Mid(pTmp, pMid, 1)

            If pChr <> "0" Then
                pCode = i

                'Exit For
            End If
            pMid = pMid - 1
        Next

        Return pCode
    End Function

    Public Function up_BitPosition(ByVal pTmp As String, ByVal pStart As Integer, ByVal pEnd As Integer) As String

        Dim pChr As String = ""
        Dim pMid As Integer = 16
        Dim pCode As String = "0"

        For i As Integer = pStart To pEnd
            pChr = Mid(pTmp, pMid, 1)

            If pChr <> "0" Then
                pCode = i

                'Exit For
            End If
            pMid = pMid - 1
        Next

        Return pCode
    End Function

    Function ReadCSV(ByVal path As String) As System.Data.DataTable

        Try

            Dim sr As New StreamReader(path)

            Dim fullFileStr As String = sr.ReadToEnd()

            sr.Close()

            sr.Dispose()

            Dim lines As String() = fullFileStr.Split(ControlChars.Lf)

            Dim recs As New DataTable()

            Dim sArr As String() = lines(0).Split(","c)

            For Each s As String In sArr

                recs.Columns.Add(New DataColumn())

            Next

            Dim row As DataRow

            Dim finalLine As String = ""

            For Each line As String In lines

                row = recs.NewRow()

                finalLine = line.Replace(Convert.ToString(ControlChars.Cr), "")

                row.ItemArray = finalLine.Split(","c)

                recs.Rows.Add(row)

            Next

            Return recs

        Catch ex As Exception

            txtMsg.Text = ex.Message

        End Try

    End Function

    Private Function uf_ProcessAlarm_Trouble(ByVal dtTmp As DataTable, ByVal pRowStart As Integer, ByVal pAlarm As String) As String

        Dim pBit_PositionPrev As Integer = up_BitPosition(pAlarm, 0, 15)
        Dim pAlarmCol_E As String = "", pAlarmCol_F As String = "", pAlarmCol_G As String = ""
        Dim pValue As String = ""

        For RowI = pRowStart To dtTmp.Rows.Count - 2

            pAlarmCol_E = up_BitPosition(DecimalToBinary(Trim(dtTmp.Rows(pRowStart)(4))), 0, 15)

            If (pBit_PositionPrev <> pAlarmCol_E) Then
                pValue = Format(CDate(20 & Trim(dtTmp.Rows(pRowStart)(0)) & " " & Trim(dtTmp.Rows(pRowStart)(1))), "yyyy-MM-dd HH:mm:ss")

                Exit For
            End If

        Next

        Return pValue

    End Function

    Private Function uf_ProcessAlarm(ByVal dtTmp As DataTable, ByVal pRowStart As Integer, ByVal pAlarm As String) As String

        Dim pBit_PositionPrev As Integer = up_BitPosition(pAlarm, 0, 15)
        Dim pAlarmCol_E As String = "", pAlarmCol_F As String = "", pAlarmCol_G As String = ""
        Dim pValue As String = ""

        For RowI = pRowStart To dtTmp.Rows.Count - 2

            pAlarmCol_E = up_BitPosition(DecimalToBinary(Trim(dtTmp.Rows(pRowStart)(4))), 0, 15)
            pAlarmCol_F = up_BitPosition(DecimalToBinary(Trim(dtTmp.Rows(pRowStart)(5))), 0, 15)
            pAlarmCol_G = up_BitPosition(DecimalToBinary(Trim(dtTmp.Rows(pRowStart)(6))), 0, 15)

            If (pBit_PositionPrev <> pAlarmCol_E) Or (pBit_PositionPrev <> pAlarmCol_F) Or (pBit_PositionPrev <> pAlarmCol_G) Then
                pValue = Format(CDate(20 & Trim(dtTmp.Rows(pRowStart)(0)) & " " & Trim(dtTmp.Rows(pRowStart)(1))), "yyyy-MM-dd HH:mm:ss")

                Exit For
            End If

        Next

        Return pValue

    End Function

    Private Function uf_ProcessAlarmInfo(ByVal dtTmp As DataTable, ByVal pRowStart As Integer, ByVal pAlarm As String) As String

        Dim pBit_PositionPrev As Integer = up_BitPosition(pAlarm, 0, 15)
        Dim pAlarmCol_E As String = "", pAlarmCol_F As String = "", pAlarmCol_G As String = ""
        Dim pValue As String = ""

        For RowI = pRowStart To dtTmp.Rows.Count - 2

            pAlarmCol_E = up_BitPosition(DecimalToBinary(Trim(dtTmp.Rows(pRowStart)(4))), 0, 15)
            pAlarmCol_F = up_BitPosition(DecimalToBinary(Trim(dtTmp.Rows(pRowStart)(5))), 0, 15)
            pAlarmCol_G = up_BitPosition(DecimalToBinary(Trim(dtTmp.Rows(pRowStart)(6))), 0, 15)

            If (pBit_PositionPrev <> pAlarmCol_E) Or (pBit_PositionPrev <> pAlarmCol_F) Or (pBit_PositionPrev <> pAlarmCol_G) Then
                pValue = Format(CDate(20 & Trim(dtTmp.Rows(pRowStart)(0)) & " " & Trim(dtTmp.Rows(pRowStart)(1))), "yyyy-MM-dd HH:mm:ss")

                Exit For
            End If

        Next

        Return pValue

    End Function

    Private Function GetFtpFileList(ByVal Url As String, ByVal userName As String, ByVal password As String, Optional ByVal fileName As String = "") As List(Of String)
        Dim FilesListNothing As New List(Of String)
        Try

            Dim request = DirectCast(WebRequest.Create(Url), FtpWebRequest)

            request.Method = WebRequestMethods.Ftp.ListDirectory

            If userName <> "OBU" And userName <> "OBO" And userName <> "OSP" Then
                request.Credentials = New NetworkCredential(userName, password)
                'Else
                '    request.Credentials = New NetworkCredential("Tos", "Tosis")
            End If

            Using reader As New StreamReader(request.GetResponse().GetResponseStream())
                Dim line = reader.ReadLine()
                Dim lines As New List(Of String)

                Do Until line Is Nothing
                    If line <> fileName Then
                        lines.Add(line)
                    End If
                    line = reader.ReadLine()
                Loop

                Return lines
            End Using

        Catch ex As Exception
            txtMsg.Text = ex.Message
            WriteToErrorLog("Get FTP File", txtMsg.Text)
            Return FilesListNothing
        End Try

    End Function

    Private Sub DownloadFiles(ByRef Files As List(Of String), ByRef Patterns As String, ByRef UserName As String, ByRef Password As String, ByRef Url As String, ByRef PathToWriteFilesTo As String,
                              ByVal pTrouble As String, ByVal pQuality As String)
        Dim Pattern() As String = Patterns.Split("|"c)
        Dim fileName As String = ""

        txtMsg.Text = ""

        Try

            For i = 0 To Files.Count - 1

                For Each Item In Pattern
                    Dim pSplit As String()
                    pSplit = Split(Files(i), "/")

                    If pSplit.Count = 1 Then
                        fileName = pSplit(0)
                    Else
                        fileName = pSplit(3)
                    End If

                    If fileName = pTrouble Or fileName = pQuality Then
                        If Files(i).ToUpper.Contains(Item.ToUpper) Then
                            'Dim request As FtpWebRequest = DirectCast(WebRequest.Create(Url & Files(i)), FtpWebRequest)
                            Dim request As FtpWebRequest = DirectCast(WebRequest.Create(Url & fileName), FtpWebRequest)

                            request.Method = WebRequestMethods.Ftp.DownloadFile ' >> Download File

                            ' This example assumes the FTP site uses anonymous logon.
                            If UserName <> "OBU" Or UserName <> "OBO" Or UserName <> "OSP" Then
                                request.Credentials = New NetworkCredential(UserName, Password)
                            End If

                            Using response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)

                                'My.Computer.FileSystem.WriteAllBytes(" ", response.GetResponseStream)
                                Using responseStream As Stream = response.GetResponseStream()
                                    Using MS As New IO.MemoryStream
                                        'MsgBox("copy")
                                        responseStream.CopyTo(MS)
                                        My.Computer.FileSystem.WriteAllBytes(PathToWriteFilesTo & "\" & fileName, MS.ToArray, False)

                                    End Using
                                End Using
                            End Using
                        End If
                    ElseIf fileName Like "*" & pTrouble & "*" Then
                        If Files(i).ToUpper.Contains(Item.ToUpper) Then
                            'Dim request As FtpWebRequest = DirectCast(WebRequest.Create(Url & Files(i)), FtpWebRequest)
                            Dim request As FtpWebRequest = DirectCast(WebRequest.Create(Url & fileName), FtpWebRequest)

                            request.Method = WebRequestMethods.Ftp.DownloadFile ' >> Download File

                            ' This example assumes the FTP site uses anonymous logon.
                            If UserName <> "OBU" And UserName <> "OBO" And UserName <> "OSP" Then
                                request.Credentials = New NetworkCredential(UserName, Password)
                                'Else
                                '    request.Credentials = New NetworkCredential("Tos", "Tosis")
                            End If

                            Using response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)

                                'My.Computer.FileSystem.WriteAllBytes(" ", response.GetResponseStream)
                                Using responseStream As Stream = response.GetResponseStream()
                                    Using MS As New IO.MemoryStream
                                        'MsgBox("copy")
                                        responseStream.CopyTo(MS)
                                        My.Computer.FileSystem.WriteAllBytes(PathToWriteFilesTo & "\" & fileName, MS.ToArray, False)

                                    End Using
                                End Using
                            End Using
                        End If
                    End If

                Next
            Next

        Catch ex As Exception
            txtMsg.Text = ex.Message
            WriteToErrorLog("Download FTP File", txtMsg.Text)
        End Try

    End Sub

    Private Sub DeleteFiles(ByRef Files As List(Of String), ByRef Patterns As String, ByRef UserName As String, ByRef Password As String, ByRef Url As String, ByRef PathToWriteFilesTo As String,
                              ByVal pTrouble As String, ByVal pQuality As String)
        Dim Pattern() As String = Patterns.Split("|"c)
        Dim fileName As String = ""

        For i = 0 To Files.Count - 1

            For Each Item In Pattern
                Dim pSplit As String()
                pSplit = Split(Files(i), "/")

                If pSplit.Count = 1 Then
                    fileName = pSplit(0)
                Else
                    fileName = pSplit(3)
                End If

                If fileName = pTrouble Or fileName = pQuality Then
                    If Files(i).ToUpper.Contains(Item.ToUpper) Then
                        'Dim request As FtpWebRequest = DirectCast(WebRequest.Create(Url & Files(i)), FtpWebRequest)
                        Dim request As FtpWebRequest = DirectCast(WebRequest.Create(Url & fileName), FtpWebRequest)

                        request.Method = WebRequestMethods.Ftp.DeleteFile

                        ' This example assumes the FTP site uses anonymous logon.
                        If UserName <> "OBU" Or UserName <> "OBO" Or UserName <> "OSP" Then
                            request.Credentials = New NetworkCredential(UserName, Password)
                        End If

                        Using response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)

                            'My.Computer.FileSystem.WriteAllBytes(" ", response.GetResponseStream)
                            Using responseStream As Stream = response.GetResponseStream()
                                Using MS As New IO.MemoryStream
                                    'MsgBox("copy")
                                    responseStream.CopyTo(MS)
                                    'My.Computer.FileSystem.WriteAllBytes(PathToWriteFilesTo & "\" & fileName, MS.ToArray, False)

                                End Using
                            End Using
                        End Using
                    End If
                ElseIf fileName Like "*" & pTrouble & "*" Then
                    If Files(i).ToUpper.Contains(Item.ToUpper) Then
                        'Dim request As FtpWebRequest = DirectCast(WebRequest.Create(Url & Files(i)), FtpWebRequest)
                        Dim request As FtpWebRequest = DirectCast(WebRequest.Create(Url & fileName), FtpWebRequest)

                        request.Method = WebRequestMethods.Ftp.DeleteFile

                        ' This example assumes the FTP site uses anonymous logon.
                        If UserName <> "OBU" And UserName <> "OBO" And UserName <> "OSP" Then
                            request.Credentials = New NetworkCredential(UserName, Password)
                            'Else
                            '    request.Credentials = New NetworkCredential("Tos", "Tosis")
                        End If

                        Using response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)

                            'My.Computer.FileSystem.WriteAllBytes(" ", response.GetResponseStream)
                            Using responseStream As Stream = response.GetResponseStream()
                                Using MS As New IO.MemoryStream
                                    'MsgBox("copy")
                                    responseStream.CopyTo(MS)
                                    'My.Computer.FileSystem.WriteAllBytes(PathToWriteFilesTo & "\" & fileName, MS.ToArray, False)

                                End Using
                            End Using
                        End Using
                    End If
                End If

            Next
        Next
    End Sub

    Private Sub TimerProcess_Tick(sender As Object, e As EventArgs) Handles TimerProcess.Tick
        up_Process()
    End Sub

    Public Sub WriteToErrorLog(ByVal pScreenName As String, Optional ByVal pErrSummary As String = "", Optional ByVal pErrID As Integer = 9999)

        Dim ls_Date As String
        Dim ls_ErrType As String
        Dim ls_dateFolder As String
        Dim ls_CompName As String = ""
        Dim ls_LogFolder As String = "D:\RelayPLC"
        ls_dateFolder = Format(Now, "yyyyMMdd")
        ls_Date = Format(Now, "yyyyMMdd")
        ls_ErrType = ""
        pScreenName = pScreenName.Trim
        pScreenName = pScreenName.Replace(" ", "_")


        'If Not System.IO.Directory.Exists("D:\") Then
        ls_LogFolder = Application.StartupPath
        'End If

        If Not System.IO.Directory.Exists(ls_LogFolder) Then
            System.IO.Directory.CreateDirectory(ls_LogFolder)
        End If

        If Not System.IO.Directory.Exists(ls_LogFolder & "\Log" &
            "\Error") Then
            System.IO.Directory.CreateDirectory(ls_LogFolder & "\Log" &
                "\Error")
        End If

        'check the file
        Dim fs As FileStream = New FileStream(ls_LogFolder & "\Log" &
            "\Error\err" & pScreenName & "_" & ls_Date & ".log", FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim s As StreamWriter = New StreamWriter(fs)

        s.Close()
        fs.Close()

        'log it
        Dim fs1 As FileStream = New FileStream(ls_LogFolder & "\Log" &
            "\Error\err" & pScreenName & "_" & ls_Date & ".log", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)

        s1.Write("" & Format(Now, "dd/MM/yyyy HH:mm:ss") & " ")
        's1.Write("[" & ls_CompName & "] ")
        s1.Write("" & pScreenName & "" & " ")
        's1.Write("[" & ls_ErrType & "]" & " ")
        s1.Write("" & pErrSummary & "" & "")
        s1.Write("" & vbCrLf)
        s1.Close()
        fs1.Close()

    End Sub

#End Region

#End Region
End Class