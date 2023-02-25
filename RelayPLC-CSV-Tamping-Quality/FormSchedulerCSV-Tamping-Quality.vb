Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Threading
Imports C1.Win.C1FlexGrid

Public Class FormSchedulerCSV_Tamping_Quality
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

    'Mixing Mach. Vacuum HT42H 
    Dim MixingMachVacuumHT42H As Integer = 1
    Dim Line_MixingMachVacuumHT42H As String = "31001"
    Dim Machine_MixingMachVacuumHT42H As String = "31001"

    'Mixing Mach. Vacuum HT35H
    Dim MixingMachVacuumHT35H As Integer = 2
    Dim Line_MixingMachVacuumHT35H As String = "31002"
    Dim Machine_MixingMachVacuumHT35H As String = "31002"

    'Mixing Mach. Vacuum EC
    Dim MixingMachVacuumEC As Integer = 3
    Dim Line_MixingMachVacuumEC As String = "31003"
    Dim Machine_MixingMachVacuumEC As String = "31003"

    'High Speed Mixer Mch
    Dim HighSpeedMixerMch As Integer = 4
    Dim Line_HighSpeedMixerMch As String = "31004"
    Dim Machine_HighSpeedMixerMch As String = "31004"

    'Conveyor transfer powder
    Dim Conveyortransferpowder As Integer = 5
    Dim Line_Conveyortransferpowder As String = "31005"
    Dim Machine_Conveyortransferpowder As String = "31005"

    'Oscilator Mach
    Dim OscilatorMach As Integer = 6
    Dim Line_OscilatorMach As String = "31006"
    Dim Machine_OscilatorMach As String = "31006"

    'Conveyor transfer box
    Dim Conveyortransferbox As Integer = 7
    Dim Line_Conveyortransferbox As String = "31007"
    Dim Machine_Conveyortransferbox As String = "31007"

    'Box lift
    Dim Boxlift As Integer = 8
    Dim Line_Boxlift As String = "31008"
    Dim Machine_Boxlift As String = "31008"

    'Expandmetal tension
    Dim Expandmetaltension As Integer = 9
    Dim Line_Expandmetaltension As String = "31009"
    Dim Machine_Expandmetaltension As String = "31009"

    'Filling
    Dim Filling As Integer = 10
    Dim Line_Filling As String = "31010"
    Dim Machine_Filling As String = "31010"

    'MOven Roller Mach. Manual #1
    Dim OvenRollerMachManual1 As Integer = 11
    Dim Line_OvenRollerMachManual1 As String = "31011"
    Dim Machine_OvenRollerMachManual1 As String = "31011"

    'Oven Heater Mach. Manual #1
    Dim OvenHeaterMachManual1 As Integer = 12
    Dim Line_OvenHeaterMachManual1 As String = "31012"
    Dim Machine_OvenHeaterMachManual1 As String = "31012"

    'Oven Roller Mach. Manual #2
    Dim OvenRollerMachManual2 As Integer = 13
    Dim Line_OvenRollerMachManual2 As String = "31013"
    Dim Machine_OvenRollerMachManual2 As String = "31013"

    'Oven Heater Mach. Manual #2
    Dim OvenHeaterMachManual2 As Integer = 14
    Dim Line_OvenHeaterMachManual2 As String = "31014"
    Dim Machine_OvenHeaterMachManual2 As String = "31014"

    'Hoop Tension
    Dim HoopTension As Integer = 15
    Dim Line_HoopTension As String = "31015"
    Dim Machine_HoopTension As String = "31015"

    'Mesin Pressing (auto manual)
    Dim MesinPressing As Integer = 16
    Dim Line_MesinPressing As String = "31016"
    Dim Machine_MesinPressing As String = "31016"

    'Roller Mach. Manual #1
    Dim RollerMachManual1 As Integer = 17
    Dim Line_RollerMachManual1 As String = "31017"
    Dim Machine_RollerMachManual1 As String = "31017"

    'Roller Mach. Manual #2
    Dim RollerMachManual2 As Integer = 18
    Dim Line_RollerMachManual2 As String = "31018"
    Dim Machine_RollerMachManual2 As String = "31018"

    'Roller Mach. Manual #3
    Dim RollerMachManual3 As Integer = 19
    Dim Line_RollerMachManual3 As String = "31019"
    Dim Machine_RollerMachManual3 As String = "31019"

    'Roller Mach. Manual #4
    Dim RollerMachManual4 As Integer = 20
    Dim Line_RollerMachManual4 As String = "31020"
    Dim Machine_RollerMachManual4 As String = "31020"

    'Roller Mach. Manual #5
    Dim RollerMachManual5 As Integer = 21
    Dim Line_RollerMachManual5 As String = "31021"
    Dim Machine_RollerMachManual5 As String = "31021"

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

    Dim Thd_MixingMachVacuumHT42H As SchedulerSetting
    Dim Thd_MixingMachVacuumHT35H As SchedulerSetting
    Dim Thd_MixingMachVacuumEC As SchedulerSetting
    Dim Thd_HighSpeedMixerMch As SchedulerSetting
    Dim Thd_Conveyortransferpowder As SchedulerSetting
    Dim Thd_OscilatorMach As SchedulerSetting
    Dim Thd_Conveyortransferbox As SchedulerSetting
    Dim Thd_Boxlift As SchedulerSetting
    Dim Thd_Expandmetaltension As SchedulerSetting
    Dim Thd_Filling As SchedulerSetting
    Dim Thd_OvenRollerMachManual1 As SchedulerSetting
    Dim Thd_OvenHeaterMachManual1 As SchedulerSetting
    Dim Thd_OvenRollerMachManual2 As SchedulerSetting
    Dim Thd_OvenHeaterMachManual2 As SchedulerSetting
    Dim Thd_HoopTension As SchedulerSetting
    Dim Thd_MesinPressing As SchedulerSetting
    Dim Thd_RollerMachManual1 As SchedulerSetting
    Dim Thd_RollerMachManual2 As SchedulerSetting
    Dim Thd_RollerMachManual3 As SchedulerSetting
    Dim Thd_RollerMachManual4 As SchedulerSetting
    Dim Thd_RollerMachManual5 As SchedulerSetting
    Dim Thd_SlitterMach1 As SchedulerSetting
    Dim Thd_SlitterMach2 As SchedulerSetting

    Dim col_Count As Integer = 6
    Dim FilesList As New List(Of String)
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

        'Mixing Mach Vacuum HT42H
        Do Until Thd_MixingMachVacuumHT42H.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_MixingMachVacuumHT42H.Status = "iddle" Then
                Thd_MixingMachVacuumHT42H.ScheduleThd = Nothing
            End If
            If Thd_MixingMachVacuumHT42H.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Mixing Mach Vacuum HT35H
        Do Until Thd_MixingMachVacuumHT35H.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_MixingMachVacuumHT35H.Status = "iddle" Then
                Thd_MixingMachVacuumHT35H.ScheduleThd = Nothing
            End If
            If Thd_MixingMachVacuumHT35H.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Mixing Mach Vacuum EC
        Do Until Thd_MixingMachVacuumEC.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_MixingMachVacuumEC.Status = "iddle" Then
                Thd_MixingMachVacuumEC.ScheduleThd = Nothing
            End If
            If Thd_MixingMachVacuumEC.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'High Speed Mixer Mch
        Do Until Thd_HighSpeedMixerMch.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_HighSpeedMixerMch.Status = "iddle" Then
                Thd_HighSpeedMixerMch.ScheduleThd = Nothing
            End If
            If Thd_HighSpeedMixerMch.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Conveyor transfer powder
        Do Until Thd_Conveyortransferpowder.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_Conveyortransferpowder.Status = "iddle" Then
                Thd_Conveyortransferpowder.ScheduleThd = Nothing
            End If
            If Thd_Conveyortransferpowder.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Oscilator Mach
        Do Until Thd_OscilatorMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_OscilatorMach.Status = "iddle" Then
                Thd_OscilatorMach.ScheduleThd = Nothing
            End If
            If Thd_OscilatorMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Conveyor transfer box
        Do Until Thd_Conveyortransferbox.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_Conveyortransferbox.Status = "iddle" Then
                Thd_Conveyortransferbox.ScheduleThd = Nothing
            End If
            If Thd_Conveyortransferbox.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Box lift
        Do Until Thd_Boxlift.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_Boxlift.Status = "iddle" Then
                Thd_Boxlift.ScheduleThd = Nothing
            End If
            If Thd_Boxlift.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Expandmetal tension
        Do Until Thd_Expandmetaltension.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_Expandmetaltension.Status = "iddle" Then
                Thd_Expandmetaltension.ScheduleThd = Nothing
            End If
            If Thd_Expandmetaltension.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Filling
        Do Until Thd_Filling.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_Filling.Status = "iddle" Then
                Thd_Filling.ScheduleThd = Nothing
            End If
            If Thd_Filling.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Oven Roller Mach Manual 1
        Do Until Thd_OvenRollerMachManual1.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_OvenRollerMachManual1.Status = "iddle" Then
                Thd_OvenRollerMachManual1.ScheduleThd = Nothing
            End If
            If Thd_OvenRollerMachManual1.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Oven Heater Mach Manual 1
        Do Until Thd_OvenHeaterMachManual1.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_OvenHeaterMachManual1.Status = "iddle" Then
                Thd_OvenHeaterMachManual1.ScheduleThd = Nothing
            End If
            If Thd_OvenHeaterMachManual1.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Oven Roller Mach Manual 2
        Do Until Thd_OvenRollerMachManual2.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_OvenRollerMachManual2.Status = "iddle" Then
                Thd_OvenRollerMachManual2.ScheduleThd = Nothing
            End If
            If Thd_OvenRollerMachManual2.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Oven Heater Mach Manual 2
        Do Until Thd_OvenHeaterMachManual2.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_OvenHeaterMachManual2.Status = "iddle" Then
                Thd_OvenHeaterMachManual2.ScheduleThd = Nothing
            End If
            If Thd_OvenHeaterMachManual2.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Hoop Tension
        Do Until Thd_HoopTension.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_HoopTension.Status = "iddle" Then
                Thd_HoopTension.ScheduleThd = Nothing
            End If
            If Thd_HoopTension.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Mesin Pressing
        Do Until Thd_MesinPressing.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_MesinPressing.Status = "iddle" Then
                Thd_MesinPressing.ScheduleThd = Nothing
            End If
            If Thd_MesinPressing.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Roller Mach Manual 1
        Do Until Thd_RollerMachManual1.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_RollerMachManual1.Status = "iddle" Then
                Thd_RollerMachManual1.ScheduleThd = Nothing
            End If
            If Thd_RollerMachManual1.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Roller Mac hManual 2
        Do Until Thd_RollerMachManual2.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_RollerMachManual2.Status = "iddle" Then
                Thd_RollerMachManual2.ScheduleThd = Nothing
            End If
            If Thd_RollerMachManual2.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Roller Mach Manual 3
        Do Until Thd_RollerMachManual3.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_RollerMachManual3.Status = "iddle" Then
                Thd_RollerMachManual3.ScheduleThd = Nothing
            End If
            If Thd_RollerMachManual3.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Roller Mach Manual 4
        Do Until Thd_RollerMachManual4.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_RollerMachManual4.Status = "iddle" Then
                Thd_RollerMachManual4.ScheduleThd = Nothing
            End If
            If Thd_RollerMachManual4.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Roller Mach Manual 5
        Do Until Thd_RollerMachManual5.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_RollerMachManual5.Status = "iddle" Then
                Thd_RollerMachManual5.ScheduleThd = Nothing
            End If
            If Thd_RollerMachManual5.ScheduleThd Is Nothing Then Exit Do
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

        up_TimeStart() 'temporay hiden modified by riana

        btnStart.Enabled = False
        btnStop.Enabled = True
        btnManual.Enabled = False
        btnConfig.Enabled = False
        btnClose.Enabled = False
    End Sub

    Private Sub timerCurr_Tick(sender As Object, e As EventArgs) Handles timerCurr.Tick
        ShowInTaskbar = True
        Me.WindowState = FormWindowState.Normal
        NotifyIcon1.Visible = False
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        lblCurrTime.Text = Format(Now, "HH:mm:ss")
        lblCurrDate.Text = Format(Now, "dddd , dd MMM yyyy")
    End Sub

    Private Sub FormSchedulerCSV_Tamping_Quality_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

    Private Sub FormSchedulerCSV_Tamping_Quality_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
                NotifyIcon1.BalloonTipTitle = "RELAY PLC CSV QUALITY (TAMPING)"
                NotifyIcon1.BalloonTipText = "Move to system tray"
                NotifyIcon1.ShowBalloonTip(50)
                ShowInTaskbar = False
            End If
        End If
    End Sub
#End Region

#Region "PROCEDURES"

    Private Sub up_GridHeader()
        With grid
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


            'Mixing Mach. Vacuum HT42H 
            '====================
            If gs_ServerPath_MixingMachVacuumHT42H <> "" And gs_LocalPath_MixingMachVacuumHT42H <> "" Then
                pProses = "High Speed Mixer Mch"
                up_ProcessData(MixingMachVacuumHT42H, "High Speed Mixer Mch", gs_LocalPath_MixingMachVacuumHT42H, gs_FileName_MixingMachVacuumHT42H, Line_MixingMachVacuumHT42H, Machine_MixingMachVacuumHT42H, Factory)
            End If
            '====================

            'Mixing Mach. Vacuum HT35H 
            '====================
            If gs_ServerPath_MixingMachVacuumHT35H <> "" And gs_LocalPath_MixingMachVacuumHT35H <> "" Then
                pProses = "Mixing Mach Vacuum HT35H"
                up_ProcessData(MixingMachVacuumHT35H, "High Speed Mixer Mch", gs_LocalPath_MixingMachVacuumHT35H, gs_FileName_MixingMachVacuumHT35H, Line_MixingMachVacuumHT35H, Machine_MixingMachVacuumHT35H, Factory)
            End If
            '====================

            'Mixing Mach. Vacuum EC 
            '====================
            If gs_ServerPath_MixingMachVacuumEC <> "" And gs_LocalPath_MixingMachVacuumEC <> "" Then
                pProses = "Mixing Mach Vacuum EC"
                up_ProcessData(MixingMachVacuumEC, "High Speed Mixer Mch", gs_LocalPath_MixingMachVacuumEC, gs_FileName_MixingMachVacuumEC, Line_MixingMachVacuumEC, Machine_MixingMachVacuumEC, Factory)
            End If
            '====================

            'High Speed Mixer Mch
            '====================
            If gs_ServerPath_HighSpeedMixerMch <> "" And gs_LocalPath_HighSpeedMixerMch <> "" Then
                pProses = "High Speed Mixer Mch"
                up_ProcessData(HighSpeedMixerMch, "High Speed Mixer Mch", gs_LocalPath_HighSpeedMixerMch, gs_FileName_HighSpeedMixerMch, Line_HighSpeedMixerMch, Machine_HighSpeedMixerMch, Factory)
            End If
            '====================

            'Conveyor transfer powder
            '====================
            If gs_ServerPath_Conveyortransferpowder <> "" And gs_LocalPath_Conveyortransferpowder <> "" Then
                pProses = "Conveyor transfer powder"
                up_ProcessData(Conveyortransferpowder, "High Speed Mixer Mch", gs_LocalPath_Conveyortransferpowder, gs_FileName_Conveyortransferpowder, Line_Conveyortransferpowder, Machine_Conveyortransferpowder, Factory)
            End If
            '====================

            'Oscilator Mach
            '====================
            If gs_ServerPath_OscilatorMach <> "" And gs_LocalPath_OscilatorMach <> "" Then
                pProses = "Oscilato Mach"
                up_ProcessData(OscilatorMach, "High Speed Mixer Mch", gs_LocalPath_OscilatorMach, gs_FileName_OscilatorMach, Line_OscilatorMach, Machine_OscilatorMach, Factory)
            End If
            '====================

            'Conveyor transfer box
            '====================
            If gs_ServerPath_Conveyortransferbox <> "" And gs_LocalPath_Conveyortransferbox <> "" Then
                pProses = "Conveyor transfer box"
                up_ProcessData(Conveyortransferbox, "High Speed Mixer Mch", gs_LocalPath_Conveyortransferbox, gs_FileName_Conveyortransferbox, Line_Conveyortransferbox, Machine_Conveyortransferbox, Factory)
            End If
            '====================

            'Box lift
            '====================
            If gs_ServerPath_Boxlift <> "" And gs_LocalPath_Boxlift <> "" Then
                pProses = "Box lift"
                up_ProcessData(Boxlift, "Box lift", gs_LocalPath_Boxlift, gs_FileName_Boxlift, Line_Boxlift, Machine_Boxlift, Factory)
            End If
            '====================

            'Expandmetal tension 
            '====================
            If gs_ServerPath_Expandmetaltension <> "" And gs_LocalPath_Expandmetaltension <> "" Then
                pProses = "Expandmetal tension"
                up_ProcessData(Expandmetaltension, "Expandmetal tension", gs_LocalPath_Expandmetaltension, gs_FileName_Expandmetaltension, Line_Expandmetaltension, Machine_Expandmetaltension, Factory)
            End If
            '====================

            'Filling
            '====================
            If gs_ServerPath_Filling <> "" And gs_LocalPath_Filling <> "" Then
                pProses = "Filling"
                up_ProcessData(Filling, "Filling", gs_LocalPath_Filling, gs_FileName_Filling, Line_Filling, Machine_Filling, Factory)
            End If
            '====================

            'Oven Roller Mach. Manual #1
            '====================
            If gs_ServerPath_OvenRollerMachManual1 <> "" And gs_LocalPath_OvenRollerMachManual1 <> "" Then
                pProses = "Oven Roller Mach Manual 1"
                up_ProcessData(OvenRollerMachManual1, "Oven Roller Mach Manual 1", gs_LocalPath_OvenRollerMachManual1, gs_FileName_OvenRollerMachManual1, Line_OvenRollerMachManual1, Machine_OvenRollerMachManual1, Factory)
            End If
            '====================

            'Oven Heater Mach. Manual #1
            '====================
            If gs_ServerPath_OvenHeaterMachManual1 <> "" And gs_LocalPath_OvenHeaterMachManual1 <> "" Then
                pProses = "Oven Heater Mach Manual 1"
                up_ProcessData(OvenHeaterMachManual1, "Oven Heater Mach Manual 1", gs_LocalPath_OvenHeaterMachManual1, gs_FileName_OvenHeaterMachManual1, Line_OvenHeaterMachManual1, Machine_OvenHeaterMachManual1, Factory)
            End If
            '====================

            'Oven Roller Mach. Manual #2
            '====================
            If gs_ServerPath_OvenRollerMachManual2 <> "" And gs_LocalPath_OvenRollerMachManual2 <> "" Then
                pProses = "Oven Roller Mach Manual 2"
                up_ProcessData(OvenRollerMachManual2, "High Speed Mixer Mch", gs_LocalPath_OvenRollerMachManual2, gs_FileName_OvenRollerMachManual2, Line_OvenRollerMachManual2, Machine_OvenRollerMachManual2, Factory)
            End If
            '====================

            'Oven Heater Mach. Manual #2
            '====================
            If gs_ServerPath_OvenHeaterMachManual2 <> "" And gs_LocalPath_OvenHeaterMachManual2 <> "" Then
                pProses = "Oven Heater Mach Manual 2"
                up_ProcessData(HighSpeedMixerMch, "Oven Heater Mach Manual 2", gs_LocalPath_OvenHeaterMachManual2, gs_FileName_OvenHeaterMachManual2, Line_OvenHeaterMachManual2, Machine_OvenHeaterMachManual2, Factory)
            End If
            '====================

            'Hoop Tension
            '====================
            If gs_ServerPath_HoopTension <> "" And gs_LocalPath_HoopTension <> "" Then
                pProses = "High Speed Mixer Mch"
                up_ProcessData(HighSpeedMixerMch, "High Speed Mixer Mch", gs_LocalPath_HighSpeedMixerMch, gs_FileName_HighSpeedMixerMch, Line_HighSpeedMixerMch, Machine_HighSpeedMixerMch, Factory)
            End If
            '====================

            'Mesin Pressing (auto manual)
            '====================
            If gs_ServerPath_MesinPressing <> "" And gs_LocalPath_MesinPressing <> "" Then
                pProses = "Mesin Pressing"
                up_ProcessData(MesinPressing, "Mesin Pressing", gs_LocalPath_MesinPressing, gs_FileName_MesinPressing, Line_MesinPressing, Machine_MesinPressing, Factory)
            End If
            '====================

            'Roller Mach. Manual #1
            '====================
            If gs_ServerPath_RollerMachManual1 <> "" And gs_LocalPath_RollerMachManual1 <> "" Then
                pProses = "Roller Mach Manual 1"
                up_ProcessData(RollerMachManual1, "Roller Mach Manual 1", gs_LocalPath_RollerMachManual1, gs_FileName_RollerMachManual1, Line_RollerMachManual1, Machine_RollerMachManual1, Factory)
            End If
            '====================

            'Roller Mach. Manual #2
            '====================
            If gs_ServerPath_RollerMachManual2 <> "" And gs_LocalPath_RollerMachManual2 <> "" Then
                pProses = "Roller Mach Manual 2"
                up_ProcessData(RollerMachManual2, "Roller Mach Manual 2", gs_LocalPath_RollerMachManual2, gs_FileName_RollerMachManual2, Line_RollerMachManual2, Machine_RollerMachManual2, Factory)
            End If
            '====================

            'Roller Mach. Manual #3
            '====================
            If gs_ServerPath_RollerMachManual3 <> "" And gs_LocalPath_RollerMachManual3 <> "" Then
                pProses = "Roller Mach Manual 3"
                up_ProcessData(RollerMachManual3, "Roller Mach Manual 3", gs_LocalPath_RollerMachManual3, gs_FileName_RollerMachManual3, Line_RollerMachManual3, Machine_RollerMachManual3, Factory)
            End If
            '====================

            'Roller Mach. Manual #4
            '====================
            If gs_ServerPath_RollerMachManual4 <> "" And gs_LocalPath_RollerMachManual4 <> "" Then
                pProses = "Roller Mach Manual 4"
                up_ProcessData(RollerMachManual4, "Roller Mach Manual 4", gs_LocalPath_RollerMachManual4, gs_FileName_RollerMachManual4, Line_RollerMachManual4, Machine_RollerMachManual4, Factory)
            End If
            '====================

            'Roller Mach. Manual #5
            '====================
            If gs_ServerPath_RollerMachManual5 <> "" And gs_LocalPath_RollerMachManual5 <> "" Then
                pProses = "Roller Mach Manual 5"
                up_ProcessData(RollerMachManual5, "Roller Mach Manual 5", gs_LocalPath_RollerMachManual5, gs_FileName_RollerMachManual5, Line_RollerMachManual5, Machine_RollerMachManual5, Factory)
            End If
            '====================

            'Slitter Mach 1
            '====================
            If gs_ServerPath_SlitterMach1 <> "" And gs_LocalPath_SlitterMach1 <> "" Then
                pProses = "Slitter Mach 1"
                up_ProcessData(SlitterMach1, "Slitter Mach 1", gs_LocalPath_SlitterMach1, gs_FileName_SlitterMach1, Line_SlitterMach1, Machine_SlitterMach1, Factory)
            End If
            '====================

            'Slitter Mach 2
            '====================
            If gs_ServerPath_SlitterMach2 <> "" And gs_LocalPath_SlitterMach2 <> "" Then
                pProses = "Slitter Mach 2"
                up_ProcessData(SlitterMach2, "Slitter Mach 2", gs_LocalPath_SlitterMach2, gs_FileName_SlitterMach2, Line_SlitterMach2, Machine_SlitterMach2, Factory)
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

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_MixingMachVacuumHT42H Then
                            gs_ServerPath_HighSpeedMixerMch = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_HighSpeedMixerMch = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_HighSpeedMixerMch = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_HighSpeedMixerMch = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_HighSpeedMixerMch = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_MixingMachVacuumHT35H Then
                            gs_ServerPath_MixingMachVacuumHT35H = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_MixingMachVacuumHT35H = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_MixingMachVacuumHT35H = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_MixingMachVacuumHT35H = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_MixingMachVacuumHT35H = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_MixingMachVacuumEC Then
                            gs_ServerPath_MixingMachVacuumEC = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_MixingMachVacuumEC = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_MixingMachVacuumEC = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_MixingMachVacuumEC = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_MixingMachVacuumEC = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_HighSpeedMixerMch Then
                            gs_ServerPath_HighSpeedMixerMch = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_HighSpeedMixerMch = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_HighSpeedMixerMch = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_HighSpeedMixerMch = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_HighSpeedMixerMch = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_Conveyortransferpowder Then
                            gs_ServerPath_Conveyortransferpowder = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_Conveyortransferpowder = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_Conveyortransferpowder = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_Conveyortransferpowder = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_Conveyortransferpowder = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_OscilatorMach Then
                            gs_ServerPath_OscilatorMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_OscilatorMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_OscilatorMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_OscilatorMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_OscilatorMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_Conveyortransferbox Then
                            gs_ServerPath_Conveyortransferbox = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_Conveyortransferbox = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_Conveyortransferbox = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_Conveyortransferbox = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_Conveyortransferbox = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_Boxlift Then
                            gs_ServerPath_Boxlift = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_Boxlift = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_Boxlift = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_Boxlift = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_Boxlift = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_Expandmetaltension Then
                            gs_ServerPath_Expandmetaltension = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_Expandmetaltension = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_Expandmetaltension = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_Expandmetaltension = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_Expandmetaltension = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Filling Then
                            gs_ServerPath_Filling = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_Filling = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_Filling = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_Filling = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_Filling = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_OvenRollerMachManual1 Then
                            gs_ServerPath_OvenRollerMachManual1 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_OvenRollerMachManual1 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_OvenRollerMachManual1 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_OvenRollerMachManual1 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_OvenRollerMachManual1 = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_OvenHeaterMachManual1 Then
                            gs_ServerPath_OvenHeaterMachManual1 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_OvenHeaterMachManual1 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_OvenHeaterMachManual1 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_OvenHeaterMachManual1 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_OvenHeaterMachManual1 = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_OvenRollerMachManual2 Then
                            gs_ServerPath_OvenRollerMachManual2 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_OvenRollerMachManual2 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_OvenRollerMachManual2 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_OvenRollerMachManual2 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_OvenRollerMachManual2 = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_OvenHeaterMachManual2 Then
                            gs_ServerPath_OvenHeaterMachManual2 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_OvenHeaterMachManual2 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_OvenHeaterMachManual2 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_OvenHeaterMachManual2 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_OvenHeaterMachManual2 = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_HoopTension Then
                            gs_ServerPath_HoopTension = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_HoopTension = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_HoopTension = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_HoopTension = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_HoopTension = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_MesinPressing Then
                            gs_ServerPath_MesinPressing = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_MesinPressing = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_MesinPressing = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_MesinPressing = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_MesinPressing = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_RollerMachManual1 Then
                            gs_ServerPath_RollerMachManual1 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_RollerMachManual1 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_RollerMachManual1 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_RollerMachManual1 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_RollerMachManual1 = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_RollerMachManual2 Then
                            gs_ServerPath_RollerMachManual2 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_RollerMachManual2 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_RollerMachManual2 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_RollerMachManual2 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_RollerMachManual2 = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_RollerMachManual3 Then
                            gs_ServerPath_RollerMachManual3 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_RollerMachManual3 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_RollerMachManual3 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_RollerMachManual3 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_RollerMachManual3 = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_RollerMachManual4 Then
                            gs_ServerPath_RollerMachManual4 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_RollerMachManual4 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_RollerMachManual4 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_RollerMachManual4 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_RollerMachManual4 = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_RollerMachManual5 Then
                            gs_ServerPath_RollerMachManual5 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_RollerMachManual5 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_RollerMachManual5 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_RollerMachManual5 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_RollerMachManual5 = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_SlitterMach1 Then
                            gs_ServerPath_SlitterMach1 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_SlitterMach1 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_SlitterMach1 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_SlitterMach1 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_SlitterMach1 = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_SlitterMach1 Then
                            gs_ServerPath_SlitterMach1 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_SlitterMach1 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_SlitterMach1 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_SlitterMach1 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_SlitterMach1 = ds.Tables(0).Rows(i)("TimeInterval")

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
            ' Mixing Mach Vacuum HT42H
            '===========================
            Thread.Sleep(200)
            Thd_MixingMachVacuumHT42H = New SchedulerSetting
            With Thd_MixingMachVacuumHT42H
                .Name = "MixingMachVacuumHT42H"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_MixingMachVacuumHT42H
                .ScheduleThd = New Thread(AddressOf up_Refresh_MixingMachVacuumHT42H)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "MixingMachVacuumHT42H"
                .ScheduleThd.Start()
            End With

            ' Mixing Mach Vacuum HT35H
            '===========================
            Thread.Sleep(200)
            Thd_MixingMachVacuumHT35H = New SchedulerSetting
            With Thd_MixingMachVacuumHT35H
                .Name = "MixingMachVacuumHT35H"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_MixingMachVacuumHT35H
                .ScheduleThd = New Thread(AddressOf up_Refresh_MixingMachVacuumHT35H)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "MixingMachVacuumHT35H"
                .ScheduleThd.Start()
            End With

            ' Mixing Mach. Vacuum EC
            '===========================
            Thread.Sleep(200)
            Thd_MixingMachVacuumEC = New SchedulerSetting
            With Thd_MixingMachVacuumEC
                .Name = "MixingMachVacuumEC"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_MixingMachVacuumEC
                .ScheduleThd = New Thread(AddressOf up_Refresh_MixingMachVacuumEC)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "MixingMachVacuumEC"
                .ScheduleThd.Start()
            End With

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

            ' Conveyor transfer powder
            '===========================
            Thread.Sleep(200)
            Thd_Conveyortransferpowder = New SchedulerSetting
            With Thd_Conveyortransferpowder
                .Name = "Conveyortransferpowder"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_Conveyortransferpowder
                .ScheduleThd = New Thread(AddressOf up_Refresh_Conveyortransferpowder)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "Conveyortransferpowder"
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

            ' Conveyor transfer box
            '===========================
            Thread.Sleep(200)
            Thd_Conveyortransferbox = New SchedulerSetting
            With Thd_Conveyortransferbox
                .Name = "Conveyortransferbox"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_Conveyortransferbox
                .ScheduleThd = New Thread(AddressOf up_Refresh_Conveyortransferbox)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "Conveyortransferbox"
                .ScheduleThd.Start()
            End With

            ' Box lift
            '===========================
            Thread.Sleep(200)
            Thd_Boxlift = New SchedulerSetting
            With Thd_Boxlift
                .Name = "Boxlift"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_Boxlift
                .ScheduleThd = New Thread(AddressOf up_Refresh_Boxlift)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "Boxlift"
                .ScheduleThd.Start()
            End With

            ' Expandmetal tension 1
            '===========================
            Thread.Sleep(200)
            Thd_Expandmetaltension = New SchedulerSetting
            With Thd_Expandmetaltension
                .Name = "Expandmetaltension"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_Expandmetaltension
                .ScheduleThd = New Thread(AddressOf up_Refresh_Expandmetaltension)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "Expandmetaltension"
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

            ' Oven Roller Mach. Manual #1
            '===========================
            Thread.Sleep(200)
            Thd_RollerMachManual1 = New SchedulerSetting
            With Thd_RollerMachManual1
                .Name = "RollerMachManual1"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_RollerMachManual1
                .ScheduleThd = New Thread(AddressOf up_Refresh_RollerMachManual1)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "RollerMachManual1"
                .ScheduleThd.Start()
            End With

            ' Oven Heater Mach. Manual #1
            '===========================
            Thread.Sleep(200)
            Thd_OvenHeaterMachManual1 = New SchedulerSetting
            With Thd_OvenHeaterMachManual1
                .Name = "OvenHeaterMachManual1"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_OvenHeaterMachManual1
                .ScheduleThd = New Thread(AddressOf up_Refresh_OvenHeaterMachManual1)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "OvenHeaterMachManual1"
                .ScheduleThd.Start()
            End With

            ' Oven Roller Mach. Manual #2
            '===========================
            Thread.Sleep(200)
            Thd_OvenRollerMachManual1 = New SchedulerSetting
            With Thd_OvenRollerMachManual1
                .Name = "OvenRollerMachManual1"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_OvenRollerMachManual1
                .ScheduleThd = New Thread(AddressOf up_Refresh_OvenRollerMachManual1)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "OvenRollerMachManual1"
                .ScheduleThd.Start()
            End With

            ' Oven Heater Mach. Manual #2
            '===========================
            Thread.Sleep(200)
            Thd_OvenHeaterMachManual2 = New SchedulerSetting
            With Thd_OvenHeaterMachManual2
                .Name = "OvenHeaterMachManual2"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_OvenHeaterMachManual2
                .ScheduleThd = New Thread(AddressOf up_Refresh_OvenHeaterMachManual2)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "OvenHeaterMachManual2"
                .ScheduleThd.Start()
            End With

            ' Hoop Tension
            '===========================
            Thread.Sleep(200)
            Thd_HoopTension = New SchedulerSetting
            With Thd_HoopTension
                .Name = "HoopTension"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_HoopTension
                .ScheduleThd = New Thread(AddressOf up_Refresh_HoopTension)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "HoopTension"
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

            ' Roller Mach. Manual #1
            '===========================
            Thread.Sleep(200)
            Thd_RollerMachManual1 = New SchedulerSetting
            With Thd_RollerMachManual1
                .Name = "RollerMachManual1"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_RollerMachManual1
                .ScheduleThd = New Thread(AddressOf up_Refresh_RollerMachManual1)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "RollerMachManual1"
                .ScheduleThd.Start()
            End With

            ' Roller Mach. Manual #2
            '===========================
            Thread.Sleep(200)
            Thd_RollerMachManual2 = New SchedulerSetting
            With Thd_RollerMachManual2
                .Name = "RollerMachManual2"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_RollerMachManual2
                .ScheduleThd = New Thread(AddressOf up_Refresh_RollerMachManual2)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "RollerMachManual2"
                .ScheduleThd.Start()
            End With

            ' Roller Mach. Manual #3
            '===========================
            Thread.Sleep(200)
            Thd_RollerMachManual3 = New SchedulerSetting
            With Thd_RollerMachManual3
                .Name = "RollerMachManual3"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_RollerMachManual3
                .ScheduleThd = New Thread(AddressOf up_Refresh_RollerMachManual3)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "RollerMachManual3"
                .ScheduleThd.Start()
            End With

            ' Roller Mach. Manual #4
            '===========================
            Thread.Sleep(200)
            Thd_RollerMachManual4 = New SchedulerSetting
            With Thd_RollerMachManual4
                .Name = "RollerMachManual4"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_RollerMachManual4
                .ScheduleThd = New Thread(AddressOf up_Refresh_RollerMachManual4)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "RollerMachManual4"
                .ScheduleThd.Start()
            End With

            ' Roller Mach. Manual #5
            '===========================
            Thread.Sleep(200)
            Thd_RollerMachManual5 = New SchedulerSetting
            With Thd_RollerMachManual5
                .Name = "RollerMachManual5"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_RollerMachManual5
                .ScheduleThd = New Thread(AddressOf up_Refresh_RollerMachManual5)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "RollerMachManual5"
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

        'Mixing Mach Vacuum HT42H
        SyncLock Thd_MixingMachVacuumHT42H.Lock
            Thd_MixingMachVacuumHT42H.EndSchedule = True
        End SyncLock

        'Mixing Mach Vacuum HT35H
        SyncLock Thd_MixingMachVacuumHT35H.Lock
            Thd_MixingMachVacuumHT35H.EndSchedule = True
        End SyncLock

        'Mixing Mach Vacuum EC
        SyncLock Thd_MixingMachVacuumEC.Lock
            Thd_MixingMachVacuumEC.EndSchedule = True
        End SyncLock

        'High Speed Mixer Mch
        SyncLock Thd_HighSpeedMixerMch.Lock
            Thd_HighSpeedMixerMch.EndSchedule = True
        End SyncLock

        'Conveyor transfer powder
        SyncLock Thd_Conveyortransferpowder.Lock
            Thd_Conveyortransferpowder.EndSchedule = True
        End SyncLock

        'Oscilator Mach
        SyncLock Thd_OscilatorMach.Lock
            Thd_OscilatorMach.EndSchedule = True
        End SyncLock

        'Conveyor transfer box
        SyncLock Thd_Conveyortransferbox.Lock
            Thd_Conveyortransferbox.EndSchedule = True
        End SyncLock

        'Box lift
        SyncLock Thd_Boxlift.Lock
            Thd_Boxlift.EndSchedule = True
        End SyncLock

        'Expandmetal tension
        SyncLock Thd_Expandmetaltension.Lock
            Thd_Expandmetaltension.EndSchedule = True
        End SyncLock

        'Filling
        SyncLock Thd_Filling.Lock
            Thd_Filling.EndSchedule = True
        End SyncLock

        'OvenRoller Mach Manual 1
        SyncLock Thd_OvenRollerMachManual1.Lock
            Thd_OvenRollerMachManual1.EndSchedule = True
        End SyncLock

        'Oven Heater Mach Manual 1
        SyncLock Thd_OvenHeaterMachManual1.Lock
            Thd_OvenHeaterMachManual1.EndSchedule = True
        End SyncLock

        'OvenRoller Mach Manual 2
        SyncLock Thd_OvenRollerMachManual2.Lock
            Thd_OvenRollerMachManual2.EndSchedule = True
        End SyncLock

        'Oven Heater Mach Manual 2
        SyncLock Thd_OvenHeaterMachManual2.Lock
            Thd_OvenHeaterMachManual2.EndSchedule = True
        End SyncLock

        'Hoop Tension
        SyncLock Thd_HoopTension.Lock
            Thd_HoopTension.EndSchedule = True
        End SyncLock

        'Mesin Pressing
        SyncLock Thd_MesinPressing.Lock
            Thd_MesinPressing.EndSchedule = True
        End SyncLock

        'Roller Mach Manual 1
        SyncLock Thd_RollerMachManual1.Lock
            Thd_RollerMachManual1.EndSchedule = True
        End SyncLock

        'RollerMachManual 2
        SyncLock Thd_RollerMachManual2.Lock
            Thd_RollerMachManual2.EndSchedule = True
        End SyncLock

        'RollerMachManual 3
        SyncLock Thd_RollerMachManual3.Lock
            Thd_RollerMachManual3.EndSchedule = True
        End SyncLock

        'RollerMachManual 4
        SyncLock Thd_RollerMachManual4.Lock
            Thd_RollerMachManual4.EndSchedule = True
        End SyncLock

        'RollerMachManual 5
        SyncLock Thd_RollerMachManual5.Lock
            Thd_MesinPressing.EndSchedule = True
        End SyncLock

        'Slitter Mach 1
        SyncLock Thd_SlitterMach1.Lock
            Thd_SlitterMach1.EndSchedule = True
        End SyncLock

        'Slitter Mach 2
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
            ls_Alarm6 As String = "0", ls_Alarm7 As String = "0", ls_Alarm8 As String = "0", ls_Alarm9 As String = "0", ls_Alarm10 As String = "0"

        'Dim pSerialNo As String = "", pSeqNo As Integer = 0, pSubSeq_No As Integer = 0
        'pSerialNo = Format(Now, "yyMMdd") & "-" & Format(Now, "HHmmss") & "-" & pLineCode

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

                    ''Save semua kolom CSV untuk Kenading saja
                    ''========================================
                    'If pLineCode = "121" Then
                    '    up_SavedataCSV(dtCSV, pLineCode, "21002")
                    'End If
                    ''========================================

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
                                    Dim alarm As String = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(3 + index))), startGroup, endGroup)

                                    If index = 1 Then
                                        ls_Alarm1 = alarm
                                    End If
                                    If index = 2 Then
                                        ls_Alarm2 = alarm
                                    End If
                                    If index = 3 Then
                                        ls_Alarm3 = alarm
                                    End If
                                    If index = 4 Then
                                        ls_Alarm4 = alarm
                                    End If
                                    If index = 5 Then
                                        ls_Alarm5 = alarm
                                    End If
                                    If index = 6 Then
                                        ls_Alarm6 = alarm
                                    End If
                                    If index = 7 Then
                                        ls_Alarm7 = alarm
                                    End If
                                    If index = 8 Then
                                        ls_Alarm8 = alarm
                                    End If
                                    If index = 9 Then
                                        ls_Alarm9 = alarm
                                    End If
                                    If index = 10 Then
                                        ls_Alarm10 = alarm
                                    End If
                                Next

                                If Trim(dtCSV.Rows(x + 1)(0)) <> "" Then
                                    EndTime = Format(CDate(20 & Trim(dtCSV.Rows(x + 1)(0)) & " " & Trim(dtCSV.Rows(x + 1)(1))), "yyyy-MM-dd HH:mm:ss")
                                Else
                                    EndTime = Format(CDate(20 & Trim(dtCSV.Rows(x)(0)) & " " & Trim(dtCSV.Rows(x)(1))), "yyyy-MM-dd HH:mm:ss")
                                End If

                                dt.Rows.Add(Col_Line, ModeCls, StatusCls, ls_Alarm1, ls_Alarm2, ls_Alarm3, ls_Alarm4, ls_Alarm5, ls_Alarm6, ls_Alarm7, ls_Alarm8, ls_Alarm9, ls_Alarm10, StartTime, EndTime, LastUpdate, "Tes")

                            End If
                        Next

                        ''Send to History Folder
                        ''======================
                        'My.Computer.FileSystem.MoveFile(pLocalPath & "\" & fi.Name, pLocalPath & "\History\" & fi.Name, True)

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

        txtMsg.Text = ""

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
                            Try
                                tmpDatde = Split(Trim(dtCSV.Rows(x)(0)), "/")
                                If tmpDatde(0) > 0 Then
                                    col_Line = pLineCode
                                    col_Machine = pMCCode
                                    col_model = "BBRSRUSA0PAD"
                                    Col_MeasTime = Format(CDate(20 & Trim(dtCSV.Rows(x)(0)) & " " & Trim(dtCSV.Rows(x)(1))), "yyyy-MM-dd HH:mm:ss")
                                    data_001 = "2" 'code from csv
                                    data_002 = Trim(dtCSV.Rows(x)(2))
                                    data_003 = Trim(dtCSV.Rows(x)(3))
                                    data_004 = Trim(dtCSV.Rows(x)(4))
                                    data_005 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(5))), 0, 15)
                                    data_006 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(6))), 0, 15)
                                    data_007 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(7))), 0, 15)
                                    data_008 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(8))), 0, 15)
                                    data_009 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(9))), 0, 15)
                                    data_010 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(10))), 0, 15)
                                    data_011 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(11))), 0, 15)
                                    data_012 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(12))), 0, 15)
                                    data_013 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(13))), 0, 15)
                                    data_014 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(14))), 0, 15)
                                    data_015 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(15))), 0, 15)
                                    data_016 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(16))), 0, 15)
                                    data_017 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(17))), 0, 15)
                                    data_018 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(18))), 0, 15)
                                    data_019 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(19))), 0, 15)
                                    data_020 = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(20))), 0, 15)
                                    data_021 = "0"
                                    data_022 = "0"
                                    data_023 = "0"
                                    data_024 = "0"
                                    data_025 = "0"

                                    dt.Rows.Add(col_Line, col_Machine, col_model, Col_MeasTime,
                                                data_001, data_002, data_003, data_004, data_005,
                                                data_006, data_007, data_008, data_009, data_010,
                                                data_011, data_012, data_013, data_014, data_015,
                                                data_016, data_017, data_018, data_019, data_020,
                                                data_021, data_022, data_023, data_024, data_025)

                                End If
                            Catch ex As Exception
                                txtMsg.Text = ex.Message
                                WriteToErrorLog("Quality Process error in LineCode : " & pLineCode, ex.Message)
                            End Try
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

    Private Function up_GetLastData(ByVal pLineCode As String) As DataTable
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim dt As New DataTable
        Dim sql As String

        Try

            Dim pThnBln As String = Format(Now, "yyMM")
            Dim TblName As String = "Trouble_" & pThnBln

            Using con = New SqlConnection(ConStr)
                con.Open()

                sql = "SELECT TOP 1 * FROM " & TblName & " WHERE LineCode = '" & pLineCode & "' ORDER BY Start_Time DESC,End_Time"

                cmd = New SqlCommand
                cmd.CommandText = sql
                cmd.Connection = con
                cmd.CommandType = CommandType.Text

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
    Private Sub up_Refresh_MixingMachVacuumHT42H()
        up_Refresh(MixingMachVacuumHT42H, "MixingMachVacuumHT42H", Thd_MixingMachVacuumHT42H, gs_ServerPath_MixingMachVacuumHT42H, gs_LocalPath_MixingMachVacuumHT42H, gs_User_MixingMachVacuumHT42H, gs_Password_MixingMachVacuumHT42H, gs_FileName_MixingMachVacuumHT42H, Line_MixingMachVacuumHT42H, Machine_MixingMachVacuumHT42H)
    End Sub

    Private Sub up_Refresh_MixingMachVacuumHT35H()
        up_Refresh(MixingMachVacuumHT42H, "MixingMachVacuumHT42H", Thd_MixingMachVacuumHT42H, gs_ServerPath_MixingMachVacuumHT42H, gs_LocalPath_MixingMachVacuumHT42H, gs_User_MixingMachVacuumHT42H, gs_Password_MixingMachVacuumHT42H, gs_FileName_MixingMachVacuumHT42H, Line_MixingMachVacuumHT42H, Machine_MixingMachVacuumHT42H)
    End Sub

    Private Sub up_Refresh_MixingMachVacuumEC()
        up_Refresh(MixingMachVacuumEC, "MixingMachVacuumEC", Thd_MixingMachVacuumEC, gs_ServerPath_MixingMachVacuumEC, gs_LocalPath_MixingMachVacuumEC, gs_User_MixingMachVacuumEC, gs_Password_MixingMachVacuumEC, gs_FileName_MixingMachVacuumEC, Line_MixingMachVacuumEC, Machine_MixingMachVacuumEC)
    End Sub

    Private Sub up_Refresh_HighSpeedMixerMch()
        up_Refresh(HighSpeedMixerMch, "HighSpeedMixerMch", Thd_HighSpeedMixerMch, gs_ServerPath_HighSpeedMixerMch, gs_LocalPath_HighSpeedMixerMch, gs_User_HighSpeedMixerMch, gs_Password_HighSpeedMixerMch, gs_FileName_HighSpeedMixerMch, Line_HighSpeedMixerMch, Machine_HighSpeedMixerMch)
    End Sub

    Private Sub up_Refresh_Conveyortransferpowder()
        up_Refresh(Conveyortransferpowder, "Conveyortransferpowder", Thd_Conveyortransferpowder, gs_ServerPath_Conveyortransferpowder, gs_LocalPath_Conveyortransferpowder, gs_User_Conveyortransferpowder, gs_Password_Conveyortransferpowder, gs_FileName_Conveyortransferpowder, Line_Conveyortransferpowder, Machine_Conveyortransferpowder)
    End Sub

    Private Sub up_Refresh_OscilatorMach()
        up_Refresh(OscilatorMach, "OscilatorMach", Thd_OscilatorMach, gs_ServerPath_OscilatorMach, gs_LocalPath_OscilatorMach, gs_User_OscilatorMach, gs_Password_OscilatorMach, gs_FileName_OscilatorMach, Line_OscilatorMach, Machine_OscilatorMach)
    End Sub

    Private Sub up_Refresh_Conveyortransferbox()
        up_Refresh(Conveyortransferbox, "Conveyortransferbox", Thd_Conveyortransferbox, gs_ServerPath_Conveyortransferbox, gs_LocalPath_Conveyortransferbox, gs_User_Conveyortransferbox, gs_Password_Conveyortransferbox, gs_FileName_Conveyortransferbox, Line_Conveyortransferbox, Machine_Conveyortransferbox)
    End Sub

    Private Sub up_Refresh_Boxlift()
        up_Refresh(Boxlift, "Boxlift", Thd_Boxlift, gs_ServerPath_Boxlift, gs_LocalPath_Boxlift, gs_User_Boxlift, gs_Password_Boxlift, gs_FileName_Boxlift, Line_Boxlift, Machine_Boxlift)
    End Sub

    Private Sub up_Refresh_Expandmetaltension()
        up_Refresh(Expandmetaltension, "Expandmetaltension", Thd_Expandmetaltension, gs_ServerPath_Expandmetaltension, gs_LocalPath_Expandmetaltension, gs_User_Expandmetaltension, gs_Password_Expandmetaltension, gs_FileName_Expandmetaltension, Line_Expandmetaltension, Machine_Expandmetaltension)
    End Sub

    Private Sub up_Refresh_Filling()
        up_Refresh(Filling, "Filling", Thd_Filling, gs_ServerPath_Filling, gs_LocalPath_Filling, gs_User_Filling, gs_Password_Filling, gs_FileName_Filling, Line_Filling, Machine_Filling)
    End Sub

    Private Sub up_Refresh_OvenRollerMachManual1()
        up_Refresh(OvenRollerMachManual1, "OvenRollerMachManual1", Thd_OvenRollerMachManual1, gs_ServerPath_OvenRollerMachManual1, gs_LocalPath_OvenRollerMachManual1, gs_User_OvenRollerMachManual1, gs_Password_OvenRollerMachManual1, gs_FileName_OvenRollerMachManual1, Line_OvenRollerMachManual1, Machine_OvenRollerMachManual1)
    End Sub

    Private Sub up_Refresh_OvenHeaterMachManual1()
        up_Refresh(OvenHeaterMachManual1, "OvenHeaterMachManual1", Thd_OvenHeaterMachManual1, gs_ServerPath_OvenHeaterMachManual1, gs_LocalPath_OvenHeaterMachManual1, gs_User_OvenHeaterMachManual1, gs_Password_OvenHeaterMachManual1, gs_FileName_OvenHeaterMachManual1, Line_OvenHeaterMachManual1, Machine_OvenHeaterMachManual1)
    End Sub

    Private Sub up_Refresh_OvenRollerMachManual2()
        up_Refresh(OvenRollerMachManual2, "OvenRollerMachManual2", Thd_OvenRollerMachManual2, gs_ServerPath_OvenRollerMachManual2, gs_LocalPath_OvenRollerMachManual2, gs_User_OvenRollerMachManual2, gs_Password_OvenRollerMachManual2, gs_FileName_OvenRollerMachManual2, Line_OvenRollerMachManual2, Machine_OvenRollerMachManual2)
    End Sub

    Private Sub up_Refresh_OvenHeaterMachManual2()
        up_Refresh(OvenHeaterMachManual2, "OvenHeaterMachManual2", Thd_OvenHeaterMachManual2, gs_ServerPath_OvenHeaterMachManual2, gs_LocalPath_OvenHeaterMachManual2, gs_User_OvenHeaterMachManual2, gs_Password_OvenHeaterMachManual2, gs_FileName_OvenHeaterMachManual2, Line_OvenHeaterMachManual2, Machine_OvenHeaterMachManual2)
    End Sub

    Private Sub up_Refresh_HoopTension()
        up_Refresh(HoopTension, "HoopTension", Thd_HoopTension, gs_ServerPath_HoopTension, gs_LocalPath_HoopTension, gs_User_HoopTension, gs_Password_HoopTension, gs_FileName_HoopTension, Line_HoopTension, Machine_HoopTension)
    End Sub

    Private Sub up_Refresh_MesinPressing()
        up_Refresh(MesinPressing, "MesinPressing", Thd_MesinPressing, gs_ServerPath_MesinPressing, gs_LocalPath_MesinPressing, gs_User_MesinPressing, gs_Password_MesinPressing, gs_FileName_MesinPressing, Line_MesinPressing, Machine_MesinPressing)
    End Sub

    Private Sub up_Refresh_RollerMachManual1()
        up_Refresh(RollerMachManual1, "RollerMachManual1", Thd_RollerMachManual1, gs_ServerPath_RollerMachManual1, gs_LocalPath_RollerMachManual1, gs_User_RollerMachManual1, gs_Password_RollerMachManual1, gs_FileName_RollerMachManual1, Line_RollerMachManual1, Machine_RollerMachManual1)
    End Sub

    Private Sub up_Refresh_RollerMachManual2()
        up_Refresh(RollerMachManual2, "RollerMachManual2", Thd_RollerMachManual2, gs_ServerPath_RollerMachManual2, gs_LocalPath_RollerMachManual2, gs_User_RollerMachManual2, gs_Password_RollerMachManual2, gs_FileName_RollerMachManual2, Line_RollerMachManual2, Machine_RollerMachManual2)
    End Sub

    Private Sub up_Refresh_RollerMachManual3()
        up_Refresh(RollerMachManual3, "RollerMachManual3", Thd_RollerMachManual3, gs_ServerPath_RollerMachManual3, gs_LocalPath_RollerMachManual3, gs_User_RollerMachManual3, gs_Password_RollerMachManual3, gs_FileName_RollerMachManual3, Line_RollerMachManual3, Machine_RollerMachManual3)
    End Sub

    Private Sub up_Refresh_RollerMachManual4()
        up_Refresh(RollerMachManual4, "RollerMachManual4", Thd_RollerMachManual4, gs_ServerPath_RollerMachManual4, gs_LocalPath_RollerMachManual4, gs_User_RollerMachManual4, gs_Password_RollerMachManual4, gs_FileName_RollerMachManual4, Line_RollerMachManual4, Machine_RollerMachManual4)
    End Sub
    Private Sub up_Refresh_RollerMachManual5()
        up_Refresh(RollerMachManual5, "RollerMachManual5", Thd_RollerMachManual5, gs_ServerPath_RollerMachManual5, gs_LocalPath_RollerMachManual5, gs_User_RollerMachManual5, gs_Password_RollerMachManual5, gs_FileName_RollerMachManual5, Line_RollerMachManual5, Machine_RollerMachManual5)
    End Sub

    Private Sub up_Refresh_SlitterMach1()
        up_Refresh(SlitterMach1, "SlitterMach1", Thd_SlitterMach1, gs_ServerPath_SlitterMach1, gs_LocalPath_SlitterMach1, gs_User_SlitterMach1, gs_Password_SlitterMach1, gs_FileName_SlitterMach1, Line_SlitterMach1, Machine_SlitterMach1)
    End Sub

    Private Sub up_Refresh_SlitterMach2()
        up_Refresh(SlitterMach2, "SlitterMach2", Thd_SlitterMach2, gs_ServerPath_SlitterMach2, gs_LocalPath_SlitterMach2, gs_User_SlitterMach2, gs_Password_SlitterMach2, gs_FileName_SlitterMach2, Line_SlitterMach2, Machine_SlitterMach2)
    End Sub

    Private Sub up_Refresh(ByVal pProcess As Integer, ByVal pProcessName As String, ByVal Thd As SchedulerSetting, ByVal pServerPath As String, ByVal pLocalPath As String, ByVal pUser As String, ByVal pPassword As String, ByVal pFileNameHistory As String, ByVal pLineCode As String, ByVal pMCCode As String, Optional pFileName As String = "")
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
                    FilesList = GetFtpFileList(pServerPath, pUser, pPassword, pFileName)
                    If FilesList.Count > 0 Then
                        'Download File
                        '=============
                        DownloadFiles(FilesList, "", pUser, pPassword, pServerPath, pLocalPath, pFileNameHistory)

                        'Delete File
                        '===========
                        DeleteFiles(FilesList, "", pUser, pPassword, pServerPath, pLocalPath, pFileNameHistory, "zz")
                    End If
                End If
                '====================

                up_ProcessData(pProcess, pProcessName, pLocalPath, pFileNameHistory, pLineCode, pMCCode, Factory)

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

    Private Sub up_ProcessData(ByVal pProcess As Integer, ByVal pProcessName As String, ByVal pLocalPath As String, ByVal pFileNameHistory As String, ByVal pLineCode As String, ByVal pMCCode As String, ByVal pFactoryCode As String)
        Dim con As New SqlConnection

        grid.AddItem("")
        grid.Item(pProcess, Col_ErrorMessage) = ""

        Dim dtInfo As New DataTable
        dtInfo = up_ToInserDatatable_Quality(pLocalPath, pFileNameHistory, pLineCode, pMCCode)

        con = New SqlConnection(ConStr)
        con.Open()

        Dim cmd As SqlCommand
        Dim SQLTrans As SqlTransaction

        SQLTrans = con.BeginTransaction

        Try
            Application.DoEvents()
            grid.Item(pProcess, Col_ProcessStatus) = "RUNNING"

            If dtInfo IsNot Nothing Then
                If dtInfo.Rows.Count > 0 Then

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

                    Dim paramFactory As New SqlParameter()
                    paramFactory.ParameterName = "@FactoryCode"
                    paramFactory.SqlDbType = SqlDbType.VarChar

                    cmd.Parameters.Add(paramTbl)
                    cmd.Parameters.Add(paramName)
                    cmd.Parameters.Add(paramFactory)

                    cmd.Parameters("@MeasData").Value = dtInfo
                    cmd.Parameters("@ProcessTime").Value = Format(Now, "yyyy-MM-dd HH:mm:ss")
                    cmd.Parameters("@FactoryCode").Value = pFactoryCode

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

        'Mixing Mach. Vacuum HT42H 
        '====================
        If gs_ServerPath_MixingMachVacuumHT42H <> "" And gs_LocalPath_MixingMachVacuumHT42H <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_MixingMachVacuumHT42H, gs_User_MixingMachVacuumHT42H, gs_Password_MixingMachVacuumHT42H)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_MixingMachVacuumHT42H, gs_Password_MixingMachVacuumHT42H, gs_ServerPath_MixingMachVacuumHT42H, gs_LocalPath_MixingMachVacuumHT42H, gs_FileName_HighSpeedMixerMch)
            End If
        End If
        '====================

        'Mixing Mach. Vacuum HT35H 
        '====================
        If gs_ServerPath_MixingMachVacuumHT35H <> "" And gs_LocalPath_MixingMachVacuumHT35H <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_MixingMachVacuumHT35H, gs_User_MixingMachVacuumHT35H, gs_Password_MixingMachVacuumHT35H)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_MixingMachVacuumHT35H, gs_Password_MixingMachVacuumHT35H, gs_ServerPath_MixingMachVacuumHT35H, gs_LocalPath_MixingMachVacuumHT35H, gs_FileName_MixingMachVacuumHT35H)
            End If
        End If
        '====================

        'Mixing Mach. Vacuum EC 
        '====================
        If gs_ServerPath_MixingMachVacuumEC <> "" And gs_LocalPath_MixingMachVacuumEC <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_MixingMachVacuumEC, gs_User_MixingMachVacuumEC, gs_Password_MixingMachVacuumEC)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_MixingMachVacuumEC, gs_Password_MixingMachVacuumEC, gs_ServerPath_MixingMachVacuumEC, gs_LocalPath_MixingMachVacuumEC, gs_FileName_MixingMachVacuumEC)
            End If
        End If
        '====================

        'High Speed Mixer Mch
        '====================
        If gs_ServerPath_HighSpeedMixerMch <> "" And gs_LocalPath_HighSpeedMixerMch <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_HighSpeedMixerMch, gs_User_HighSpeedMixerMch, gs_Password_HighSpeedMixerMch)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_HighSpeedMixerMch, gs_Password_HighSpeedMixerMch, gs_ServerPath_HighSpeedMixerMch, gs_LocalPath_HighSpeedMixerMch, gs_FileName_HighSpeedMixerMch)
            End If
        End If
        '====================

        'Conveyor transfer powder
        '====================
        If gs_ServerPath_Conveyortransferpowder <> "" And gs_LocalPath_Conveyortransferpowder <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_Conveyortransferpowder, gs_User_Conveyortransferpowder, gs_Password_Conveyortransferpowder)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_Conveyortransferpowder, gs_Password_Conveyortransferpowder, gs_ServerPath_Conveyortransferpowder, gs_LocalPath_Conveyortransferpowder, gs_FileName_Conveyortransferpowder)
            End If
        End If
        '====================

        'Oscilator Mach
        '====================
        If gs_ServerPath_OscilatorMach <> "" And gs_LocalPath_OscilatorMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_OscilatorMach, gs_User_OscilatorMach, gs_Password_OscilatorMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_OscilatorMach, gs_Password_OscilatorMach, gs_ServerPath_OscilatorMach, gs_LocalPath_OscilatorMach, gs_FileName_OscilatorMach)
            End If
        End If
        '====================

        'Conveyor transfer box
        '====================
        If gs_ServerPath_Conveyortransferbox <> "" And gs_LocalPath_Conveyortransferbox <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_Conveyortransferbox, gs_User_Conveyortransferbox, gs_Password_Conveyortransferbox)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_Conveyortransferbox, gs_Password_Conveyortransferbox, gs_ServerPath_Conveyortransferbox, gs_LocalPath_Conveyortransferbox, gs_FileName_Conveyortransferbox)
            End If
        End If
        '====================

        'Box lift
        '====================
        If gs_ServerPath_Boxlift <> "" And gs_LocalPath_Boxlift <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_Boxlift, gs_User_Boxlift, gs_Password_Boxlift)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_Boxlift, gs_Password_Boxlift, gs_ServerPath_Boxlift, gs_LocalPath_Boxlift, gs_FileName_Boxlift)
            End If
        End If
        '====================

        'Expandmetal tension 1
        '====================
        If gs_ServerPath_Expandmetaltension <> "" And gs_LocalPath_Expandmetaltension <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_Expandmetaltension, gs_User_Expandmetaltension, gs_Password_Expandmetaltension)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_Expandmetaltension, gs_Password_Expandmetaltension, gs_ServerPath_Expandmetaltension, gs_LocalPath_Expandmetaltension, gs_FileName_Expandmetaltension)
            End If
        End If
        '====================

        'Filling
        '====================
        If gs_ServerPath_Filling <> "" And gs_LocalPath_Filling <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_Filling, gs_User_Filling, gs_Password_Filling)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_Filling, gs_Password_Filling, gs_ServerPath_Filling, gs_LocalPath_Filling, gs_FileName_Filling)
            End If
        End If
        '====================

        'Oven Roller Mach. Manual #1
        '====================
        If gs_ServerPath_OvenRollerMachManual1 <> "" And gs_LocalPath_OvenRollerMachManual1 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_OvenRollerMachManual1, gs_User_OvenRollerMachManual1, gs_Password_OvenRollerMachManual1)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_OvenRollerMachManual1, gs_Password_OvenRollerMachManual1, gs_ServerPath_OvenRollerMachManual1, gs_LocalPath_OvenRollerMachManual1, gs_FileName_OvenRollerMachManual1)
            End If
        End If
        '====================

        'Oven Heater Mach. Manual #1
        '====================
        If gs_ServerPath_OvenHeaterMachManual1 <> "" And gs_LocalPath_OvenHeaterMachManual1 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_OvenHeaterMachManual1, gs_User_OvenHeaterMachManual1, gs_Password_OvenHeaterMachManual1)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_OvenHeaterMachManual1, gs_Password_OvenHeaterMachManual1, gs_ServerPath_OvenHeaterMachManual1, gs_LocalPath_OvenHeaterMachManual1, gs_FileName_OvenHeaterMachManual1)
            End If
        End If
        '====================

        'Oven Roller Mach. Manual #2
        '====================
        If gs_ServerPath_OvenRollerMachManual2 <> "" And gs_LocalPath_OvenRollerMachManual2 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_OvenRollerMachManual2, gs_User_OvenRollerMachManual2, gs_Password_OvenRollerMachManual2)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_OvenRollerMachManual2, gs_Password_OvenRollerMachManual2, gs_ServerPath_OvenRollerMachManual2, gs_LocalPath_OvenRollerMachManual2, gs_FileName_OvenRollerMachManual2)
            End If
        End If
        '====================

        'Oven Heater Mach. Manual #2
        '====================
        If gs_ServerPath_OvenHeaterMachManual2 <> "" And gs_LocalPath_OvenHeaterMachManual2 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_OvenHeaterMachManual2, gs_User_OvenHeaterMachManual2, gs_Password_OvenHeaterMachManual2)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_OvenHeaterMachManual2, gs_Password_OvenHeaterMachManual2, gs_ServerPath_OvenHeaterMachManual2, gs_LocalPath_OvenHeaterMachManual2, gs_FileName_OvenHeaterMachManual2)
            End If
        End If
        '====================

        'Hoop Tension
        '====================
        If gs_ServerPath_HoopTension <> "" And gs_LocalPath_HoopTension <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_HoopTension, gs_User_HoopTension, gs_Password_HoopTension)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_HoopTension, gs_Password_HoopTension, gs_ServerPath_HoopTension, gs_LocalPath_HoopTension, gs_FileName_HoopTension)
            End If
        End If
        '====================

        'Mesin Pressing (auto manual)
        '====================
        If gs_ServerPath_MesinPressing <> "" And gs_LocalPath_MesinPressing <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_MesinPressing, gs_User_MesinPressing, gs_Password_MesinPressing)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_MesinPressing, gs_Password_MesinPressing, gs_ServerPath_MesinPressing, gs_LocalPath_MesinPressing, gs_FileName_MesinPressing)
            End If
        End If
        '====================

        'Roller Mach. Manual #1
        '====================
        If gs_ServerPath_RollerMachManual1 <> "" And gs_LocalPath_RollerMachManual1 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_RollerMachManual1, gs_User_RollerMachManual1, gs_Password_RollerMachManual1)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_RollerMachManual1, gs_Password_RollerMachManual1, gs_ServerPath_RollerMachManual1, gs_LocalPath_RollerMachManual1, gs_FileName_RollerMachManual1)
            End If
        End If
        '====================

        'Roller Mach. Manual #2
        '====================
        If gs_ServerPath_RollerMachManual2 <> "" And gs_LocalPath_RollerMachManual2 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_RollerMachManual2, gs_User_RollerMachManual2, gs_Password_RollerMachManual2)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_RollerMachManual2, gs_Password_RollerMachManual2, gs_ServerPath_RollerMachManual2, gs_LocalPath_RollerMachManual2, gs_FileName_RollerMachManual2)
            End If
        End If
        '====================

        'Roller Mach. Manual #3
        '====================
        If gs_ServerPath_RollerMachManual3 <> "" And gs_LocalPath_RollerMachManual3 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_RollerMachManual3, gs_User_RollerMachManual3, gs_Password_RollerMachManual3)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_RollerMachManual3, gs_Password_RollerMachManual3, gs_ServerPath_RollerMachManual3, gs_LocalPath_RollerMachManual3, gs_FileName_RollerMachManual3)
            End If
        End If
        '====================

        'Roller Mach. Manual #4
        '====================
        If gs_ServerPath_RollerMachManual4 <> "" And gs_LocalPath_RollerMachManual4 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_RollerMachManual4, gs_User_RollerMachManual4, gs_Password_RollerMachManual4)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_RollerMachManual4, gs_Password_RollerMachManual4, gs_ServerPath_RollerMachManual4, gs_LocalPath_RollerMachManual4, gs_FileName_RollerMachManual4)
            End If
        End If
        '====================

        'Roller Mach. Manual #5
        '====================
        If gs_ServerPath_RollerMachManual5 <> "" And gs_LocalPath_RollerMachManual5 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_RollerMachManual5, gs_User_RollerMachManual5, gs_Password_RollerMachManual5)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_RollerMachManual5, gs_Password_RollerMachManual5, gs_ServerPath_RollerMachManual5, gs_LocalPath_RollerMachManual5, gs_FileName_RollerMachManual5)
            End If
        End If
        '====================

        'Slitter Mach 1
        '====================
        If gs_ServerPath_SlitterMach1 <> "" And gs_LocalPath_SlitterMach1 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_SlitterMach1, gs_User_SlitterMach1, gs_Password_SlitterMach1)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_SlitterMach1, gs_Password_SlitterMach1, gs_ServerPath_SlitterMach1, gs_LocalPath_SlitterMach1, gs_FileName_SlitterMach1)
            End If
        End If
        '====================

        'Slitter Mach 2
        '====================
        If gs_ServerPath_SlitterMach2 <> "" And gs_LocalPath_SlitterMach2 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_SlitterMach2, gs_User_SlitterMach2, gs_Password_SlitterMach2)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_SlitterMach2, gs_Password_SlitterMach2, gs_ServerPath_SlitterMach2, gs_LocalPath_SlitterMach2, gs_FileName_SlitterMach2)
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
            trouble = DecimalToBinary(Trim(dtCSV.Rows(idx)(groupCount + 3)))
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

    Private Sub DownloadFiles(ByRef Files As List(Of String), ByRef Patterns As String, ByRef UserName As String, ByRef Password As String, ByRef Url As String, ByRef PathToWriteFilesTo As String, ByVal pQuality As String)
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

                    If fileName = pQuality Then
                        If Files(i).ToUpper.Contains(Item.ToUpper) Then
                            Dim request As FtpWebRequest = DirectCast(WebRequest.Create(Url & fileName), FtpWebRequest)
                            request.Method = WebRequestMethods.Ftp.DownloadFile ' Download File

                            ' This example assumes the FTP site uses anonymous logon.
                            If UserName <> "OBU" Or UserName <> "OBO" Or UserName <> "OSP" Then
                                request.Credentials = New NetworkCredential(UserName, Password)
                            End If

                            Using response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)

                                Using responseStream As Stream = response.GetResponseStream()
                                    Using MS As New MemoryStream
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