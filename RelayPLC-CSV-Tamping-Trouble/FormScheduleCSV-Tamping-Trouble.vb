﻿Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Threading
Imports C1.Win.C1FlexGrid

Public Class FormScheduleCSV_Tamping_Trouble
#Region "INITIAL"

    Private cConfig As clsConfig

    Const ConnectionErrorMsg As String = "A network-related or instance-specific error occurred while establishing a connection to SQL Server"
    Const TransportErrorMsg As String = "A transport-level error has occurred"

    Public Sub New()
        InitializeComponent()
    End Sub

#End Region

#Region "DECLARATION"

    Dim Factory As String = "F004"
    Dim Group As Integer = 3

    'High Speed Mixer Mch
    Dim HighSpeedMixerMch As Integer = 4
    Dim Line_HighSpeedMixerMch As String = "31004"
    Dim Machine_HighSpeedMixerMch As String = "31004"

    'Oscilator Mach
    Dim OscilatorMach As Integer = 6
    Dim Line_OscilatorMach As String = "31006"
    Dim Machine_OscilatorMach As String = "31006"

    'Filling
    Dim Filling As Integer = 10
    Dim Line_Filling As String = "31010"
    Dim Machine_Filling As String = "31010"

    'Mesin Pressing (auto manual)
    Dim MesinPressing As Integer = 16
    Dim Line_MesinPressing As String = "31016"
    Dim Machine_MesinPressing As String = "31016"

    'Slitter Mach 1
    Dim SlitterMach1 As Integer = 22
    Dim Line_SlitterMach1 As String = "31022"
    Dim Machine_SlitterMach1 As String = "31022"

    'Slitter Mach 2
    Dim SlitterMach2 As Integer = 23
    Dim Line_SlitterMach2 As String = "31023"
    Dim Machine_SlitterMach2 As String = "31023"

    Dim col_ProcessName As Integer = 0
    Dim Col_ProcessType As Byte = 1
    Dim Col_ProcessStatus As Byte = 2
    Dim Col_LastProcess As Byte = 3
    Dim Col_NextProcess As Byte = 4
    Dim Col_ErrorMessage As Byte = 5

    Dim col_Count As Integer = 6

    Dim FilesList As New List(Of String)


    Dim Thd_HighSpeedMixerMch As SchedulerSetting
    Dim Thd_OscilatorMach As SchedulerSetting
    Dim Thd_Filling As SchedulerSetting
    Dim Thd_MesinPressing As SchedulerSetting
    Dim Thd_SlitterMach1 As SchedulerSetting
    Dim Thd_SlitterMach2 As SchedulerSetting

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
        col_V = 21
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

        Do Until Thd_HighSpeedMixerMch.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_HighSpeedMixerMch.Status = "iddle" Then
                Thd_HighSpeedMixerMch.ScheduleThd = Nothing
            End If
            If Thd_HighSpeedMixerMch.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        Do Until Thd_OscilatorMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_OscilatorMach.Status = "iddle" Then
                Thd_OscilatorMach.ScheduleThd = Nothing
            End If
            If Thd_OscilatorMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        Do Until Thd_Filling.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_Filling.Status = "iddle" Then
                Thd_Filling.ScheduleThd = Nothing
            End If
            If Thd_Filling.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        Do Until Thd_MesinPressing.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_MesinPressing.Status = "iddle" Then
                Thd_MesinPressing.ScheduleThd = Nothing
            End If
            If Thd_MesinPressing.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        Do Until Thd_SlitterMach1.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_SlitterMach1.Status = "iddle" Then
                Thd_SlitterMach1.ScheduleThd = Nothing
            End If
            If Thd_SlitterMach1.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        Do Until Thd_SlitterMach2.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_SlitterMach2.Status = "iddle" Then
                Thd_SlitterMach2.ScheduleThd = Nothing
            End If
            If Thd_SlitterMach2.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
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

    Private Sub FormScheduleCSV_Tamping_Trouble_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

    Private Sub FormScheduleCSV_Tamping_Trouble_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                'NotifyIcon1.Icon = SystemIcons.Application
                NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
                NotifyIcon1.BalloonTipTitle = "RELAY PLC CSV TROUBLE (TAMPING)"
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

                sql = "SP_FTPTampingLoad_Grid"

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

            'High Speed Mixer Mch
            '====================
            If gs_ServerPath_HighSpeedMixerMch <> "" And gs_LocalPath_HighSpeedMixerMch <> "" Then
                pProses = "High Speed Mixe rMch"
                up_ProcessData(HighSpeedMixerMch, "High Speed Mixer Mch", gs_LocalPath_HighSpeedMixerMch, gs_FileName_History_HighSpeedMixerMch, Line_HighSpeedMixerMch, Group)
            End If
            '====================

            'Oscilator Mach
            '====================
            If gs_ServerPath_OscilatorMach <> "" And gs_LocalPath_OscilatorMach <> "" Then
                pProses = "Oscilato Mach"
                up_ProcessData(OscilatorMach, "High Speed Mixer Mch", gs_LocalPath_OscilatorMach, gs_FileName_History_OscilatorMach, Line_OscilatorMach, Group)
            End If
            '====================

            'Mesin Pressing (auto manual)
            '====================
            If gs_ServerPath_MesinPressing <> "" And gs_LocalPath_MesinPressing <> "" Then
                pProses = "Mesin Pressing"
                up_ProcessData(MesinPressing, "Mesin Pressing", gs_LocalPath_MesinPressing, gs_FileName_History_MesinPressing, Line_MesinPressing, Group)
            End If
            '====================


            'Slitter Mach 1
            '====================
            If gs_ServerPath_SlitterMach1 <> "" And gs_LocalPath_SlitterMach1 <> "" Then
                pProses = "Slitter Mach 1"
                up_ProcessData(SlitterMach1, "Slitter Mach 1", gs_LocalPath_SlitterMach1, gs_FileName_History_SlitterMach1, Line_SlitterMach1, Group)
            End If
            '====================

            'Slitter Mach 2
            '====================
            If gs_ServerPath_SlitterMach2 <> "" And gs_LocalPath_SlitterMach2 <> "" Then
                pProses = "Slitter Mach 2"
                up_ProcessData(SlitterMach2, "Slitter Mach 2", gs_LocalPath_SlitterMach2, gs_FileName_History_SlitterMach2, Line_SlitterMach2, Group)
            End If
            '====================

        Catch ex As Exception
            WriteToErrorLog(pProses, "error : " & ex.Message)
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

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_HighSpeedMixerMch Then
                            gs_ServerPath_HighSpeedMixerMch = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_HighSpeedMixerMch = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_HighSpeedMixerMch = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_HighSpeedMixerMch = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_HighSpeedMixerMch = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                    Next

                End If

            End Using

        Catch ex As Exception
            WriteToErrorLog("Load data FTP error", ex.Message)
        End Try

    End Sub

    Private Sub up_ClearVariable()

        gs_ServerPath_HighSpeedMixerMch = ""
        gs_LocalPath_HighSpeedMixerMch = ""
        gs_User_HighSpeedMixerMch = ""
        gs_Password_HighSpeedMixerMch = ""

    End Sub

    Private Sub up_TimeStart()
        m_Finish = False
        Me.Cursor = Cursors.WaitCursor

        Try
            ' High Speed Mixer Mch
            '===========================
            Thread.Sleep(200)
            Thd_HighSpeedMixerMch = New SchedulerSetting
            With Thd_HighSpeedMixerMch
                .Name = "HighSpeedMixerMch"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_HighSpeedMixerMch
                .ScheduleThd = New Thread(AddressOf up_Refresh_HighSpeedMixerMch)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "HighSpeedMixerMch"
                .ScheduleThd.Start()
            End With

            ' Oscilator Mach
            '===========================
            Thread.Sleep(200)
            Thd_OscilatorMach = New SchedulerSetting
            With Thd_OscilatorMach
                .Name = "OscilatorMach"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_OscilatorMach
                .ScheduleThd = New Thread(AddressOf up_Refresh_OscilatorMach)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "OscilatorMach"
                .ScheduleThd.Start()
            End With

            ' Expandmetal tension 2
            '===========================
            Thread.Sleep(200)
            Thd_Filling = New SchedulerSetting
            With Thd_Filling
                .Name = "Filling"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_Filling
                .ScheduleThd = New Thread(AddressOf up_Refresh_Filling)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "Filling"
                .ScheduleThd.Start()
            End With

            ' Mesin Pressing (auto manual)
            '===========================
            Thread.Sleep(200)
            Thd_MesinPressing = New SchedulerSetting
            With Thd_MesinPressing
                .Name = "MesinPressing"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_MesinPressing
                .ScheduleThd = New Thread(AddressOf up_Refresh_MesinPressing)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "MesinPressing"
                .ScheduleThd.Start()
            End With

            ' Slitter Mach 1
            '===========================
            Thread.Sleep(200)
            Thd_SlitterMach1 = New SchedulerSetting
            With Thd_SlitterMach1
                .Name = "SlitterMach1"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_SlitterMach1
                .ScheduleThd = New Thread(AddressOf up_Refresh_SlitterMach1)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "SlitterMach1"
                .ScheduleThd.Start()
            End With

            ' Slitter Mach 2
            '===========================
            Thread.Sleep(200)
            Thd_SlitterMach2 = New SchedulerSetting
            With Thd_SlitterMach2
                .Name = "SlitterMach2"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_SlitterMach2
                .ScheduleThd = New Thread(AddressOf up_Refresh_SlitterMach2)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "SlitterMach2"
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

        SyncLock Thd_HighSpeedMixerMch.Lock
            Thd_HighSpeedMixerMch.EndSchedule = True
        End SyncLock

        SyncLock Thd_OscilatorMach.Lock
            Thd_OscilatorMach.EndSchedule = True
        End SyncLock

        SyncLock Thd_Filling.Lock
            Thd_Filling.EndSchedule = True
        End SyncLock
        SyncLock Thd_MesinPressing.Lock
            Thd_MesinPressing.EndSchedule = True
        End SyncLock

        SyncLock Thd_SlitterMach1.Lock
            Thd_SlitterMach1.EndSchedule = True
        End SyncLock

        SyncLock Thd_SlitterMach2.Lock
            Thd_SlitterMach2.EndSchedule = True
        End SyncLock
    End Sub

    Private Function up_ToInserDatatable_Trouble(ByVal pLocalPath As String, ByVal pFileName As String, ByVal pLineCode As String, ByVal pGroupCount As Integer) As DataTable
        Dim dt As New DataTable
        Dim tmpDatde As String()
        Dim Col_Line As String = ""
        Dim pAlarm1 As String = "", EndTime As String, LastUpdate As String
        Dim pMachine As String = "", StartTime As String = "", ModeCls As String = "", StatusCls As String = ""
        Dim li_Add As Integer = 0, pBitValue As String = "", pChr As String = ""
        Dim pMid As Integer = 0
        Dim ls_Alarm1 As String = "0", ls_Alarm2 As String = "0", ls_Alarm3 As String = "0", ls_Alarm4 As String = "0", ls_Alarm5 As String = "0",
            ls_Alarm6 As String = "0", ls_Alarm7 As String = "0", ls_Alarm8 As String = "0", ls_Alarm9 As String = "0", ls_Alarm10 As String = "0",
            ls_Alarm11 As String = "0", ls_Alarm12 As String = "0", ls_Alarm13 As String = "0", ls_Alarm14 As String = "0"

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
                            .Add("TroubleCode1", GetType(String))
                            .Add("TroubleCode2", GetType(String))
                            .Add("TroubleCode3", GetType(String))
                            .Add("TroubleCode4", GetType(String))
                            .Add("TroubleCode5", GetType(String))
                            .Add("TroubleCode6", GetType(String))
                            .Add("TroubleCode7", GetType(String))
                            .Add("TroubleCode8", GetType(String))
                            .Add("TroubleCode9", GetType(String))
                            .Add("TroubleCode10", GetType(String))
                            .Add("TroubleCode11", GetType(String))
                            .Add("TroubleCode12", GetType(String))
                            .Add("TroubleCode13", GetType(String))
                            .Add("TroubleCode14", GetType(String))
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

                                For index = 1 To pGroupCount
                                    Dim startGroup As Int16 = Convert.ToInt16(list(index - 1).Split(":")(0))
                                    Dim endGroup As Int16 = Convert.ToInt16(list(index - 1).Split(":")(1))
                                    Dim alarm As String = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(5 + index))), startGroup, endGroup)

                                    If index = 1 Then
                                        ls_Alarm1 = alarm
                                    ElseIf index = 2 Then
                                        ls_Alarm2 = alarm
                                    ElseIf index = 3 Then
                                        ls_Alarm3 = alarm
                                    End If
                                Next

                                ls_Alarm4 = Trim(dtCSV.Rows(x)(4))
                                ls_Alarm5 = Trim(dtCSV.Rows(x)(5))
                                ls_Alarm6 = Trim(dtCSV.Rows(x)(6))
                                ls_Alarm7 = Trim(dtCSV.Rows(x)(7))
                                ls_Alarm8 = Trim(dtCSV.Rows(x)(8))
                                ls_Alarm9 = Trim(dtCSV.Rows(x)(9))
                                ls_Alarm10 = Trim(dtCSV.Rows(x)(10))
                                ls_Alarm11 = Trim(dtCSV.Rows(x)(11))
                                ls_Alarm12 = Trim(dtCSV.Rows(x)(12))
                                ls_Alarm13 = Trim(dtCSV.Rows(x)(13))
                                ls_Alarm14 = Trim(dtCSV.Rows(x)(14))

                                If Trim(dtCSV.Rows(x + 1)(0)) <> "" Then
                                    EndTime = Format(CDate(20 & Trim(dtCSV.Rows(x + 1)(0)) & " " & Trim(dtCSV.Rows(x + 1)(1))), "yyyy-MM-dd HH:mm:ss")
                                Else
                                    EndTime = Format(CDate(20 & Trim(dtCSV.Rows(x)(0)) & " " & Trim(dtCSV.Rows(x)(1))), "yyyy-MM-dd HH:mm:ss")
                                End If

                                dt.Rows.Add(Col_Line, ModeCls, StatusCls, ls_Alarm1, ls_Alarm2, ls_Alarm3, ls_Alarm4, ls_Alarm5, ls_Alarm6, ls_Alarm7, ls_Alarm8, ls_Alarm9, ls_Alarm10, ls_Alarm11, ls_Alarm12, ls_Alarm13, ls_Alarm14, StartTime, EndTime, LastUpdate, "Tes")

                            End If
                        Next

                        Return dt

                    End If

                End If

            Next

        Catch ex As Exception
            WriteToErrorLog("Trouble Process error in LineCode : " & pLineCode, ex.Message)
        End Try

    End Function

    Private Function up_ToInserDatatable_Trouble2(ByVal pLocalPath As String, ByVal pFileName As String, ByVal pLineCode As String, ByVal pGroupTroubleCount As Integer) As DataTable
        Dim dt As New DataTable
        Dim tmpDatde As String()
        Dim Col_Line As String = "", AlarmCode As String
        Dim pAlarm1 As String = "", EndTime As String, LastUpdate As String
        Dim pMachine As String = "", StartTime As String = "", ModeCls As String = "", StatusCls As String = ""
        Dim li_Add As Integer = 0, ls_Trouble As Boolean = False
        Dim ls_Alarm1 As String = "", ls_Alarm2 As String = "", ls_Alarm3 As String = ""

        Dim tmp_StartTime As String = "", tmp_Alarm1 As String = "", tmp_Alarm2 As String = "", tmp_Alarm3 As String = ""

        Try

            Dim di As New IO.DirectoryInfo(pLocalPath)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
            Dim fi As IO.FileInfo = Nothing

            Dim x As Integer
            Dim strarray(1, 1) As String
            Dim pCls As String = ""
            Dim iRow_Data As Integer = 0

            For Each fi In aryFi

                If fi.Name = pFileName Then

                    Dim dtCSV As DataTable = ReadCSV(fi.FullName)
                    Dim dt_LastData As DataTable = up_GetLastData(pLineCode)

                    'Backup CSV on database
                    up_SavedataCSV(dtCSV, pLineCode)

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

                        If dt_LastData.Rows.Count > 0 Then
                            tmp_Alarm1 = Trim(dt_LastData.Rows(0).Item("TroubleCode"))
                            tmp_Alarm2 = Trim(dt_LastData.Rows(0).Item("TroubleCode"))
                            tmp_Alarm3 = Trim(dt_LastData.Rows(0).Item("TroubleCode"))

                            If Trim(dt_LastData.Rows(0).Item("TroubleCode")) = "0" Then
                                ls_Trouble = False
                            Else
                                ls_Trouble = True
                            End If
                        End If

                        For x = 0 To dtCSV.Rows.Count - 2
                            tmpDatde = Split(Trim(dtCSV.Rows(x)(0)), "/")
                            Col_Line = pLineCode

                            If tmpDatde(0) > 0 Then

                                StartTime = Format(CDate(20 & Trim(dtCSV.Rows(x)(0)) & " " & Trim(dtCSV.Rows(x)(1))), "yyyy-MM-dd HH:mm:ss")
                                ModeCls = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(2))), 0, 15)
                                StatusCls = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(3))), 0, 15)
                                LastUpdate = Format(Now, "yyyy-MM-dd HH:mm:ss")

                                If dt_LastData.Rows.Count > 0 Then
                                    If Trim(dt_LastData.Rows(0).Item("TroubleCode")) <> "0" Then
                                        ls_Trouble = True
                                    End If
                                End If

                                Dim startGroup As Int16, endGroup As Int16
                                Dim group As Int16

                                group = GetGroup(dtCSV, x, pGroupTroubleCount) 'dapatkan group yg memilik nika
                                EndTime = ""

                                If (group <> 0) Then

                                    startGroup = Convert.ToInt16(list(group - 1).Split(":")(0))
                                    endGroup = Convert.ToInt16(list(group - 1).Split(":")(1))

                                    If x = 0 Then 'Row Pertama otomatis loging
                                        AlarmCode = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(group + 5))), startGroup, endGroup)
                                        dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                        iRow_Data = iRow_Data + 1
                                    Else
                                        If (tmp_Alarm3 <> ls_Alarm3 And ls_Trouble = False) Then
                                            EndTime = StartTime
                                            dt.Rows(iRow_Data - 1)("EndTime") = StartTime
                                            AlarmCode = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(group + 5))), startGroup, endGroup)
                                            dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                            iRow_Data = iRow_Data + 1

                                            ls_Trouble = True

                                        End If
                                    End If
                                Else
                                    EndTime = ""

                                    If x = 0 Then 'Row Pertama otomatis loging
                                        startGroup = Convert.ToInt16(list(pGroupTroubleCount - 1).Split(":")(0))
                                        endGroup = Convert.ToInt16(list(pGroupTroubleCount - 1).Split(":")(1))

                                        AlarmCode = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(pGroupTroubleCount + 5))), startGroup, endGroup)
                                        dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                        iRow_Data = iRow_Data + 1
                                    Else

                                        If tmp_Alarm3 <> ls_Alarm3 Then
                                            EndTime = StartTime
                                            dt.Rows(iRow_Data - 1)("EndTime") = StartTime
                                            AlarmCode = 0
                                            dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                            iRow_Data = iRow_Data + 1
                                            ls_Trouble = False

                                        End If
                                    End If
                                End If

                                tmp_StartTime = StartTime
                                tmp_Alarm1 = ls_Alarm1
                                tmp_Alarm2 = ls_Alarm2
                                tmp_Alarm3 = AlarmCode

                                'End If

                            End If
                        Next

                        Return dt

                    End If

                End If

            Next

        Catch ex As Exception
            WriteToErrorLog("Trouble Process 2 error in LineCode : " & pLineCode, ex.Message)
        End Try

    End Function

    Private Function up_GetLastData(ByVal pLineCode As String) As DataTable
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim dt As New DataTable

        Try

            Dim pThnBln As String = Format(Now, "yyMM")
            Dim TblName As String = "Trouble_" & pThnBln

            Using con = New SqlConnection(ConStr)
                con.Open()
                cmd = New SqlCommand("sp_CSVTrouble_CreateTable", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Connection = con
                cmd.Parameters.AddWithValue("TableName", TblName)
                cmd.Parameters.AddWithValue("LineCode", pLineCode)

                da = New SqlDataAdapter(cmd)
                da.Fill(dt)
                cmd.Parameters.Clear()
                cmd.Dispose()
            End Using

            Return dt

        Catch ex As Exception
            txtMsg.Text = ex.Message
        End Try

    End Function

    Private Sub up_Refresh_HighSpeedMixerMch()
        up_Refresh(HighSpeedMixerMch, "HighSpeedMixerMch", Thd_HighSpeedMixerMch, gs_ServerPath_HighSpeedMixerMch, gs_LocalPath_HighSpeedMixerMch, gs_User_HighSpeedMixerMch, gs_Password_HighSpeedMixerMch, gs_FileName_History_HighSpeedMixerMch, Line_HighSpeedMixerMch, Group)
    End Sub

    Private Sub up_Refresh_OscilatorMach()
        up_Refresh(OscilatorMach, "OscilatorMach", Thd_OscilatorMach, gs_ServerPath_OscilatorMach, gs_LocalPath_OscilatorMach, gs_User_OscilatorMach, gs_Password_OscilatorMach, gs_FileName_History_OscilatorMach, Line_OscilatorMach, Group)
    End Sub

    Private Sub up_Refresh_Filling()
        up_Refresh(Filling, "Filling", Thd_Filling, gs_ServerPath_Filling, gs_LocalPath_Filling, gs_User_Filling, gs_Password_Filling, gs_FileName_History_Filling, Line_Filling, Group)
    End Sub

    Private Sub up_Refresh_MesinPressing()
        up_Refresh(MesinPressing, "MesinPressing", Thd_MesinPressing, gs_ServerPath_MesinPressing, gs_LocalPath_MesinPressing, gs_User_MesinPressing, gs_Password_MesinPressing, gs_FileName_History_MesinPressing, Line_MesinPressing, Group)
    End Sub

    Private Sub up_Refresh_SlitterMach1()
        up_Refresh(SlitterMach1, "SlitterMach1", Thd_SlitterMach1, gs_ServerPath_SlitterMach1, gs_LocalPath_SlitterMach1, gs_User_SlitterMach1, gs_Password_SlitterMach1, gs_FileName_History_SlitterMach1, Line_SlitterMach1, Group)
    End Sub

    Private Sub up_Refresh_SlitterMach2()
        up_Refresh(SlitterMach2, "SlitterMach2", Thd_SlitterMach2, gs_ServerPath_SlitterMach2, gs_LocalPath_SlitterMach2, gs_User_SlitterMach2, gs_Password_SlitterMach2, gs_FileName_History_SlitterMach2, Line_SlitterMach2, Group)
    End Sub

    Private Sub up_Refresh(ByVal pProcess As Integer, ByVal pProcessName As String, ByVal Thd As SchedulerSetting, ByVal pServerPath As String, ByVal pLocalPath As String, ByVal pUser As String, ByVal pPassword As String, ByVal pFileNameHistory As String, ByVal pLineCode As String, ByVal pGroupCount As Integer)
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
                        DownloadFiles(FilesList, "", pUser, pPassword, pServerPath, pLocalPath, pFileNameHistory)

                        'Delete File
                        '===========
                        'DeleteFiles(FilesList, "", pUser, pPassword, pServerPath, pLocalPath, pFileNameHistory, "zz")
                    End If
                End If
                '====================

                up_ProcessData(pProcess, pProcessName, pLocalPath, pFileNameHistory, pGroupCount, pLineCode)

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

    Private Sub up_ProcessData(ByVal pProcess As Integer, ByVal pProcessName As String, ByVal pLocalPath As String, ByVal pFileNameHistory As String, ByVal pLineCode As String, ByVal pGroupCount As Integer)
        Dim con As New SqlConnection
        'gs_LocalPathMIX = "D:\PECGI CSV\Tamping\LOG_Mc21001"

        grid.AddItem("")
        grid.Item(pProcess, Col_ErrorMessage) = ""

        Dim dtHis, dtHis2 As New DataTable
        dtHis = up_ToInserDatatable_Trouble(pLocalPath, pFileNameHistory, pLineCode, pGroupCount)
        dtHis2 = up_ToInserDatatable_Trouble2(pLocalPath, pFileNameHistory, pLineCode, pGroupCount)


        Dim cmd As SqlCommand
        Dim SQLTrans As SqlTransaction
        con = New SqlConnection(ConStr)
        con.Open()

        SQLTrans = con.BeginTransaction

        Try

            Application.DoEvents()
            grid.Item(pProcess, Col_ProcessStatus) = "RUNNING"

            If dtHis IsNot Nothing Then
                If dtHis.Rows.Count > 0 Then

                    cmd = New SqlCommand("sp_Insert_Log_CSV_Trouble", con)
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

            If dtHis2 IsNot Nothing Then
                If dtHis2.Rows.Count > 0 Then

                    'clsSchedulerCSV_BA_DB.InsertData_Info(dtMixInfo, "MIX_Info_CSV_")

                    cmd = New SqlCommand("sp_Insert_Meas_CSV_Trouble", con)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Connection = con
                    cmd.Transaction = SQLTrans

                    Dim paramTbl As New SqlParameter()
                    paramTbl.ParameterName = "@MeasDataTrouble"
                    paramTbl.SqlDbType = SqlDbType.Structured

                    Dim paramName As New SqlParameter()
                    paramName.ParameterName = "@ProcessTime"
                    paramName.SqlDbType = SqlDbType.DateTime

                    cmd.Parameters.Add(paramTbl)
                    cmd.Parameters.Add(paramName)

                    cmd.Parameters("@MeasDataTrouble").Value = dtHis2
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


        'High Speed Mixer Mch
        '====================
        If gs_ServerPath_HighSpeedMixerMch <> "" And gs_LocalPath_HighSpeedMixerMch <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_HighSpeedMixerMch, gs_User_HighSpeedMixerMch, gs_Password_HighSpeedMixerMch)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_HighSpeedMixerMch, gs_Password_HighSpeedMixerMch, gs_ServerPath_HighSpeedMixerMch, gs_LocalPath_HighSpeedMixerMch, gs_FileName_History_HighSpeedMixerMch)
            End If
        End If
        '====================

        'Oscilator Mach
        '====================
        If gs_ServerPath_OscilatorMach <> "" And gs_LocalPath_OscilatorMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_OscilatorMach, gs_User_OscilatorMach, gs_Password_OscilatorMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_OscilatorMach, gs_Password_OscilatorMach, gs_ServerPath_OscilatorMach, gs_LocalPath_OscilatorMach, gs_FileName_History_OscilatorMach)
            End If
        End If
        '====================

        'Mesin Pressing (auto manual)
        '====================
        If gs_ServerPath_MesinPressing <> "" And gs_LocalPath_MesinPressing <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_MesinPressing, gs_User_MesinPressing, gs_Password_MesinPressing)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_MesinPressing, gs_Password_MesinPressing, gs_ServerPath_MesinPressing, gs_LocalPath_MesinPressing, gs_FileName_History_MesinPressing)
            End If
        End If
        '====================

        'Slitter Mach 1
        '====================
        If gs_ServerPath_SlitterMach1 <> "" And gs_LocalPath_SlitterMach1 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_SlitterMach1, gs_User_SlitterMach1, gs_Password_SlitterMach1)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_SlitterMach1, gs_Password_SlitterMach1, gs_ServerPath_SlitterMach1, gs_LocalPath_SlitterMach1, gs_FileName_History_SlitterMach1)
            End If
        End If
        '====================

        'Slitter Mach 2
        '====================
        If gs_ServerPath_SlitterMach2 <> "" And gs_LocalPath_SlitterMach2 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_SlitterMach2, gs_User_SlitterMach2, gs_Password_SlitterMach2)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_SlitterMach2, gs_Password_SlitterMach2, gs_ServerPath_SlitterMach2, gs_LocalPath_SlitterMach2, gs_FileName_History_SlitterMach2)
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
        End With

        For x = 0 To dtTmp.Rows.Count - 1
            If Trim(dtTmp.Rows(x).Item(0)) <> "" Then
                dt.Rows.Add(Format(CDate(20 & (Trim(dtTmp.Rows(x).Item(0)))), "yyyy-MM-dd"), Format(CDate(dtTmp.Rows(x).Item(1)), "HH:mm:ss"), Trim(dtTmp.Rows(x).Item(2)), Trim(dtTmp.Rows(x).Item(3)),
                            Trim(dtTmp.Rows(x).Item(4)), Trim(dtTmp.Rows(x).Item(5)), Trim(dtTmp.Rows(x).Item(6)), Trim(dtTmp.Rows(x).Item(7)), Trim(dtTmp.Rows(x).Item(8)), Trim(dtTmp.Rows(x).Item(9)),
                            Trim(dtTmp.Rows(x).Item(10)), Trim(dtTmp.Rows(x).Item(11)), Trim(dtTmp.Rows(x).Item(12)), Trim(dtTmp.Rows(x).Item(13)), Trim(dtTmp.Rows(x).Item(14)), Trim(dtTmp.Rows(x).Item(15)))
            End If
        Next

        con = New SqlConnection(ConStr)
        con.Open()

        Dim cmd As SqlCommand
        Dim SQLTrans As SqlTransaction

        SQLTrans = con.BeginTransaction

        Try

            cmd = New SqlCommand("sp_PLC_CSV_INFO", con)
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
        Dim iCount As Integer = 0 + pStart

        For i As Integer = pStart To pEnd
            pChr = Mid(pTmp, pMid, 1)
            iCount = iCount + 1
            If pChr <> "0" Then
                pCode = iCount

                'Exit For
            End If
            pMid = pMid - 1
        Next

        Return pCode
    End Function

    Public Function up_GetLastBit(ByVal pTmp As String, ByVal pStart As Integer, ByVal pEnd As Integer) As Integer

        Dim pChr As String = ""
        Dim pMid As Integer = 16
        Dim pCode As Integer = 0

        For i As Integer = pStart To pEnd
            pChr = Mid(pTmp, pMid, 1)

            If pChr = "1" Then
                pCode = i
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

    Function GetGroup(ByVal dtCSV As System.Data.DataTable, ByVal idx As Int16, ByVal groupCount As Int16) As Integer

        Do While groupCount > 0
            Dim trouble As String
            trouble = DecimalToBinary(Trim(dtCSV.Rows(idx)(groupCount + 5)))
            If (trouble <> 0) Then
                Return groupCount
            Else
                groupCount = groupCount - 1
            End If
        Loop
        Return 0
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
                              ByVal pTrouble As String)
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

                    If fileName = pTrouble Then
                        If Files(i).ToUpper.Contains(Item.ToUpper) Then
                            Dim request As FtpWebRequest = DirectCast(WebRequest.Create(Url & fileName), FtpWebRequest)
                            request.Method = WebRequestMethods.Ftp.DownloadFile ' >> Download File

                            ' This example assumes the FTP site uses anonymous logon.
                            If UserName <> "OBU" Or UserName <> "OBO" Or UserName <> "OSP" Then
                                request.Credentials = New NetworkCredential(UserName, Password)
                            End If

                            Using response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)
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

End Class
