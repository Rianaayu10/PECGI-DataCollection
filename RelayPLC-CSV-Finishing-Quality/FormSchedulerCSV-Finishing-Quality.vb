Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Threading
Imports C1.Win.C1FlexGrid

Public Class FormSchedulerCSV_Finishing_Quality
#Region "INITIAL"
    Private cConfig As clsConfig
    Const ConnectionErrorMsg As String = "A network-related or instance-specific error occurred while establishing a connection to SQL Server"
    Const TransportErrorMsg As String = "A transport-level error has occurred"

    Public Sub New()
        InitializeComponent()
    End Sub

#End Region

#Region "DECLARATION"

    Dim Factory As String = "F003"
    Dim Group As Integer = 3

    'CAP CHECKER Mac 
    Dim CAPCHECKERMac As Integer = 1
    Dim Line_CAPCHECKERMac As String = "34010"
    Dim Machine_CAPCHECKERMac As String = "34010"

    'CELL SUPPLY BATTERY Mach
    Dim CELLSUPPLYBATTERYMach As Integer = 2
    Dim Line_CELLSUPPLYBATTERYMach As String = "34011"
    Dim Machine_CELLSUPPLYBATTERYMach As String = "34011"

    'RING INSERT Mach
    Dim RINGINSERTMach As Integer = 3
    Dim Line_RINGINSERTMach As String = "34012"
    Dim Machine_RINGINSERTMach As String = "34012"

    'TUBE INSERT Mach
    Dim TUBEINSERTMach As Integer = 4
    Dim Line_TUBEINSERTMach As String = "34013"
    Dim Machine_TUBEINSERTMach As String = "34013"

    'TUBE SHRINK Mach
    Dim TUBESHRINKMach As Integer = 5
    Dim Line_TUBESHRINKMach As String = "34014"
    Dim Machine_TUBESHRINKMach As String = "34014"

    'SIDE CHECKER Mach
    Dim SIDECHECKERMach As Integer = 6
    Dim Line_SIDECHECKERMach As String = "34015"
    Dim Machine_SIDECHECKERMach As String = "34015"

    'CELL BOXING Mach
    Dim CELLBOXINGMach As Integer = 7
    Dim Line_CELLBOXINGMach As String = "34016"
    Dim Machine_CELLBOXINGMach As String = "31016"

    'SUPPLY BOX MACH Mach
    Dim SUPPLYBOXMACHMach As Integer = 8
    Dim Line_SUPPLYBOXMACHMach As String = "34017"
    Dim Machine_SUPPLYBOXMACHMach As String = "34017"

    'CELL SUPPLY Mach
    Dim CELLSUPPLYMach As Integer = 9
    Dim Line_CELLSUPPLYMach As String = "34018"
    Dim Machine_CELLSUPPLYMach As String = "34018"

    'UPPER RING INSERT Mach
    Dim UPPERRINGINSERTMach As Integer = 10
    Dim Line_UPPERRINGINSERTMach As String = "34019"
    Dim Machine_UPPERRINGINSERTMach As String = "34019"

    'LABELLER Mach
    Dim LABELLERMach As Integer = 11
    Dim Line_LABELLERMach As String = "34020"
    Dim Machine_LABELLERMach As String = "34020"

    'OCV INSPECTION Mach
    Dim OCVINSPECTIONMach As Integer = 12
    Dim Line_OCVINSPECTIONMach As String = "34021"
    Dim Machine_OCVINSPECTIONMach As String = "34021"

    'SIDE INSPECTION Mach
    Dim SIDEINSPECTIONMach As Integer = 13
    Dim Line_SIDEINSPECTIONMach As String = "34022"
    Dim Machine_SIDEINSPECTIONMach As String = "34022"

    'CELL BOXING Mach 2
    Dim CELLBOXINGMach2 As Integer = 14
    Dim Line_CELLBOXINGMach2 As String = "31014"
    Dim Machine_CELLBOXINGMach2 As String = "31014"


    Dim col_ProcessName As Integer = 0
    Dim Col_ProcessType As Byte = 1
    Dim Col_ProcessStatus As Byte = 2
    Dim Col_LastProcess As Byte = 3
    Dim Col_NextProcess As Byte = 4
    Dim Col_ErrorMessage As Byte = 5

    Dim Thd_CAPCHECKERMac As SchedulerSetting
    Dim Thd_CELLSUPPLYBATTERYMach As SchedulerSetting
    Dim Thd_RINGINSERTMach As SchedulerSetting
    Dim Thd_TUBEINSERTMach As SchedulerSetting
    Dim Thd_TUBESHRINKMach As SchedulerSetting
    Dim Thd_SIDECHECKERMach As SchedulerSetting
    Dim Thd_CELLBOXINGMach As SchedulerSetting
    Dim Thd_SUPPLYBOXMACHMach As SchedulerSetting
    Dim Thd_CELLSUPPLYMach As SchedulerSetting
    Dim Thd_UPPERRINGINSERTMach As SchedulerSetting
    Dim Thd_LABELLERMach As SchedulerSetting
    Dim Thd_OCVINSPECTIONMach As SchedulerSetting
    Dim Thd_SIDEINSPECTIONMach As SchedulerSetting
    Dim Thd_CELLBOXINGMach2 As SchedulerSetting


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
        Do Until Thd_CAPCHECKERMac.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_CAPCHECKERMac.Status = "iddle" Then
                Thd_CAPCHECKERMac.ScheduleThd = Nothing
            End If
            If Thd_CAPCHECKERMac.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Mixing Mach Vacuum HT35H
        Do Until Thd_CELLSUPPLYBATTERYMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_CELLSUPPLYBATTERYMach.Status = "iddle" Then
                Thd_CELLSUPPLYBATTERYMach.ScheduleThd = Nothing
            End If
            If Thd_CELLSUPPLYBATTERYMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Mixing Mach Vacuum EC
        Do Until Thd_RINGINSERTMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_RINGINSERTMach.Status = "iddle" Then
                Thd_RINGINSERTMach.ScheduleThd = Nothing
            End If
            If Thd_RINGINSERTMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'High Speed Mixer Mch
        Do Until Thd_TUBEINSERTMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_TUBEINSERTMach.Status = "iddle" Then
                Thd_TUBEINSERTMach.ScheduleThd = Nothing
            End If
            If Thd_TUBEINSERTMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Conveyor transfer powder
        Do Until Thd_TUBESHRINKMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_TUBESHRINKMach.Status = "iddle" Then
                Thd_TUBESHRINKMach.ScheduleThd = Nothing
            End If
            If Thd_TUBESHRINKMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Oscilator Mach
        Do Until Thd_SIDECHECKERMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_SIDECHECKERMach.Status = "iddle" Then
                Thd_SIDECHECKERMach.ScheduleThd = Nothing
            End If
            If Thd_SIDECHECKERMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Conveyor transfer box
        Do Until Thd_CELLBOXINGMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_CELLBOXINGMach.Status = "iddle" Then
                Thd_CELLBOXINGMach.ScheduleThd = Nothing
            End If
            If Thd_CELLBOXINGMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Box lift
        Do Until Thd_SUPPLYBOXMACHMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_SUPPLYBOXMACHMach.Status = "iddle" Then
                Thd_SUPPLYBOXMACHMach.ScheduleThd = Nothing
            End If
            If Thd_SUPPLYBOXMACHMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Expandmetal tension
        Do Until Thd_CELLSUPPLYMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_CELLSUPPLYMach.Status = "iddle" Then
                Thd_CELLSUPPLYMach.ScheduleThd = Nothing
            End If
            If Thd_CELLSUPPLYMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'UPPERRINGINSERTMach
        Do Until Thd_UPPERRINGINSERTMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_UPPERRINGINSERTMach.Status = "iddle" Then
                Thd_UPPERRINGINSERTMach.ScheduleThd = Nothing
            End If
            If Thd_UPPERRINGINSERTMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Oven Roller Mach Manual 1
        Do Until Thd_LABELLERMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_LABELLERMach.Status = "iddle" Then
                Thd_LABELLERMach.ScheduleThd = Nothing
            End If
            If Thd_LABELLERMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Oven Heater Mach Manual 1
        Do Until Thd_OCVINSPECTIONMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_OCVINSPECTIONMach.Status = "iddle" Then
                Thd_OCVINSPECTIONMach.ScheduleThd = Nothing
            End If
            If Thd_OCVINSPECTIONMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Oven Roller Mach Manual 2
        Do Until Thd_SIDEINSPECTIONMach.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_SIDEINSPECTIONMach.Status = "iddle" Then
                Thd_SIDEINSPECTIONMach.ScheduleThd = Nothing
            End If
            If Thd_SIDEINSPECTIONMach.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        'Oven Heater Mach Manual 2
        Do Until Thd_CELLBOXINGMach2.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_CELLBOXINGMach2.Status = "iddle" Then
                Thd_CELLBOXINGMach2.ScheduleThd = Nothing
            End If
            If Thd_CELLBOXINGMach2.ScheduleThd Is Nothing Then Exit Do
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

    Private Sub FormSchedulerCSV_Finishing_Quality_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

    Private Sub FormSchedulerCSV_Finishing_Quality_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
                NotifyIcon1.BalloonTipTitle = "RELAY PLC CSV QUALITY (Finishing)"
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

                sql = "SP_FTPFinishingLoad_Grid"

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
            If gs_ServerPath_CAPCHECKERMac <> "" And gs_LocalPath_CAPCHECKERMac <> "" Then
                pProses = "High Speed Mixer Mch"
                up_ProcessData(CAPCHECKERMac, "High Speed Mixer Mch", gs_LocalPath_CAPCHECKERMac, gs_FileName_CAPCHECKERMac, Line_CAPCHECKERMac, Machine_CAPCHECKERMac, Factory)
            End If
            '====================

            'Mixing Mach. Vacuum HT35H 
            '====================
            If gs_ServerPath_CELLSUPPLYBATTERYMach <> "" And gs_LocalPath_CELLSUPPLYBATTERYMach <> "" Then
                pProses = "Mixing Mach Vacuum HT35H"
                up_ProcessData(CELLSUPPLYBATTERYMach, "High Speed Mixer Mch", gs_LocalPath_CELLSUPPLYBATTERYMach, gs_FileName_CELLSUPPLYBATTERYMach, Line_CELLSUPPLYBATTERYMach, Machine_CELLSUPPLYBATTERYMach, Factory)
            End If
            '====================

            'Mixing Mach. Vacuum EC 
            '====================
            If gs_ServerPath_RINGINSERTMach <> "" And gs_LocalPath_RINGINSERTMach <> "" Then
                pProses = "Mixing Mach Vacuum EC"
                up_ProcessData(RINGINSERTMach, "High Speed Mixer Mch", gs_LocalPath_RINGINSERTMach, gs_FileName_RINGINSERTMach, Line_RINGINSERTMach, Machine_RINGINSERTMach, Factory)
            End If
            '====================

            'High Speed Mixer Mch
            '====================
            If gs_ServerPath_TUBEINSERTMach <> "" And gs_LocalPath_TUBEINSERTMach <> "" Then
                pProses = "High Speed Mixer Mch"
                up_ProcessData(TUBEINSERTMach, "High Speed Mixer Mch", gs_LocalPath_TUBEINSERTMach, gs_FileName_TUBEINSERTMach, Line_TUBEINSERTMach, Machine_TUBEINSERTMach, Factory)
            End If
            '====================

            'Conveyor transfer powder
            '====================
            If gs_ServerPath_TUBESHRINKMach <> "" And gs_LocalPath_TUBESHRINKMach <> "" Then
                pProses = "Conveyor transfer powder"
                up_ProcessData(TUBESHRINKMach, "High Speed Mixer Mch", gs_LocalPath_TUBESHRINKMach, gs_FileName_TUBESHRINKMach, Line_TUBESHRINKMach, Machine_TUBESHRINKMach, Factory)
            End If
            '====================

            'Oscilator Mach
            '====================
            If gs_ServerPath_SIDECHECKERMach <> "" And gs_LocalPath_SIDECHECKERMach <> "" Then
                pProses = "Oscilato Mach"
                up_ProcessData(SIDECHECKERMach, "High Speed Mixer Mch", gs_LocalPath_SIDECHECKERMach, gs_FileName_SIDECHECKERMach, Line_SIDECHECKERMach, Machine_SIDECHECKERMach, Factory)
            End If
            '====================

            'Conveyor transfer box
            '====================
            If gs_ServerPath_CELLBOXINGMach <> "" And gs_LocalPath_CELLBOXINGMach <> "" Then
                pProses = "Conveyor transfer box"
                up_ProcessData(CELLBOXINGMach, "High Speed Mixer Mch", gs_LocalPath_CELLBOXINGMach, gs_FileName_CELLBOXINGMach, Line_CELLBOXINGMach, Machine_CELLBOXINGMach, Factory)
            End If
            '====================

            'Box lift
            '====================
            If gs_ServerPath_SUPPLYBOXMACHMach <> "" And gs_LocalPath_SUPPLYBOXMACHMach <> "" Then
                pProses = "Box lift"
                up_ProcessData(SUPPLYBOXMACHMach, "Box lift", gs_LocalPath_SUPPLYBOXMACHMach, gs_FileName_SUPPLYBOXMACHMach, Line_SUPPLYBOXMACHMach, Machine_SUPPLYBOXMACHMach, Factory)
            End If
            '====================

            'Expandmetal tension 
            '====================
            If gs_ServerPath_CELLSUPPLYMach <> "" And gs_LocalPath_CELLSUPPLYMach <> "" Then
                pProses = "Expandmetal tension"
                up_ProcessData(CELLSUPPLYMach, "Expandmetal tension", gs_LocalPath_CELLSUPPLYMach, gs_FileName_CELLSUPPLYMach, Line_CELLSUPPLYMach, Machine_CELLSUPPLYMach, Factory)
            End If
            '====================

            'UPPERRINGINSERTMach
            '====================
            If gs_ServerPath_UPPERRINGINSERTMach <> "" And gs_LocalPath_UPPERRINGINSERTMach <> "" Then
                pProses = "UPPERRINGINSERTMach"
                up_ProcessData(UPPERRINGINSERTMach, "UPPERRINGINSERTMach", gs_LocalPath_UPPERRINGINSERTMach, gs_FileName_UPPERRINGINSERTMach, Line_UPPERRINGINSERTMach, Machine_UPPERRINGINSERTMach, Factory)
            End If
            '====================

            'Oven Roller Mach. Manual #1
            '====================
            If gs_ServerPath_LABELLERMach <> "" And gs_LocalPath_LABELLERMach <> "" Then
                pProses = "Oven Roller Mach Manual 1"
                up_ProcessData(LABELLERMach, "Oven Roller Mach Manual 1", gs_LocalPath_LABELLERMach, gs_FileName_LABELLERMach, Line_LABELLERMach, Machine_LABELLERMach, Factory)
            End If
            '====================

            'Oven Heater Mach. Manual #1
            '====================
            If gs_ServerPath_OCVINSPECTIONMach <> "" And gs_LocalPath_OCVINSPECTIONMach <> "" Then
                pProses = "Oven Heater Mach Manual 1"
                up_ProcessData(OCVINSPECTIONMach, "Oven Heater Mach Manual 1", gs_LocalPath_OCVINSPECTIONMach, gs_FileName_OCVINSPECTIONMach, Line_OCVINSPECTIONMach, Machine_OCVINSPECTIONMach, Factory)
            End If
            '====================

            'Oven Roller Mach. Manual #2
            '====================
            If gs_ServerPath_SIDEINSPECTIONMach <> "" And gs_LocalPath_SIDEINSPECTIONMach <> "" Then
                pProses = "Oven Roller Mach Manual 2"
                up_ProcessData(SIDEINSPECTIONMach, "High Speed Mixer Mch", gs_LocalPath_SIDEINSPECTIONMach, gs_FileName_SIDEINSPECTIONMach, Line_SIDEINSPECTIONMach, Machine_SIDEINSPECTIONMach, Factory)
            End If
            '====================

            'Oven Heater Mach. Manual #2
            '====================
            If gs_ServerPath_CELLBOXINGMach2 <> "" And gs_LocalPath_CELLBOXINGMach2 <> "" Then
                pProses = "Oven Heater Mach Manual 2"
                up_ProcessData(CELLBOXINGMach2, "Oven Heater Mach Manual 2", gs_LocalPath_CELLBOXINGMach2, gs_FileName_CELLBOXINGMach2, Line_CELLBOXINGMach2, Machine_CELLBOXINGMach2, Factory)
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

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_CAPCHECKERMac Then
                            gs_ServerPath_CAPCHECKERMac = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_CAPCHECKERMac = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_CAPCHECKERMac = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_CAPCHECKERMac = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_CAPCHECKERMac = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_CELLSUPPLYBATTERYMach Then
                            gs_ServerPath_CELLSUPPLYBATTERYMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_CELLSUPPLYBATTERYMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_CELLSUPPLYBATTERYMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_CELLSUPPLYBATTERYMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_CELLSUPPLYBATTERYMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_RINGINSERTMach Then
                            gs_ServerPath_RINGINSERTMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_RINGINSERTMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_RINGINSERTMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_RINGINSERTMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_RINGINSERTMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_TUBEINSERTMach Then
                            gs_ServerPath_TUBEINSERTMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_TUBEINSERTMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_TUBEINSERTMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_TUBEINSERTMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_TUBEINSERTMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_TUBESHRINKMach Then
                            gs_ServerPath_TUBESHRINKMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_TUBESHRINKMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_TUBESHRINKMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_TUBESHRINKMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_TUBESHRINKMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_SIDECHECKERMach Then
                            gs_ServerPath_SIDECHECKERMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_SIDECHECKERMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_SIDECHECKERMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_SIDECHECKERMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_SIDECHECKERMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_CELLBOXINGMach Then
                            gs_ServerPath_CELLBOXINGMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_CELLBOXINGMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_CELLBOXINGMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_CELLBOXINGMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_CELLBOXINGMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_SUPPLYBOXMACHMach Then
                            gs_ServerPath_SUPPLYBOXMACHMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_SUPPLYBOXMACHMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_SUPPLYBOXMACHMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_SUPPLYBOXMACHMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_SUPPLYBOXMACHMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_CELLSUPPLYMach Then
                            gs_ServerPath_CELLSUPPLYMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_CELLSUPPLYMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_CELLSUPPLYMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_CELLSUPPLYMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_CELLSUPPLYMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = UPPERRINGINSERTMach Then
                            gs_ServerPath_UPPERRINGINSERTMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_UPPERRINGINSERTMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_UPPERRINGINSERTMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_UPPERRINGINSERTMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_UPPERRINGINSERTMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_LABELLERMach Then
                            gs_ServerPath_LABELLERMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_LABELLERMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_LABELLERMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_LABELLERMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_LABELLERMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_OCVINSPECTIONMach Then
                            gs_ServerPath_OCVINSPECTIONMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_OCVINSPECTIONMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_OCVINSPECTIONMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_OCVINSPECTIONMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_OCVINSPECTIONMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_SIDEINSPECTIONMach Then
                            gs_ServerPath_SIDEINSPECTIONMach = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_SIDEINSPECTIONMach = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_SIDEINSPECTIONMach = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_SIDEINSPECTIONMach = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_SIDEINSPECTIONMach = ds.Tables(0).Rows(i)("TimeInterval")

                        ElseIf ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = Machine_CELLBOXINGMach2 Then
                            gs_ServerPath_CELLBOXINGMach2 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPath_CELLBOXINGMach2 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_User_CELLBOXINGMach2 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_Password_CELLBOXINGMach2 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_Interval_CELLBOXINGMach2 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                    Next

                End If

            End Using

        Catch ex As Exception
            WriteToErrorLog("Load data FTP error", ex.Message)
        End Try

    End Sub

    Private Sub up_ClearVariable()

        gs_ServerPath_TUBEINSERTMach = ""
        gs_LocalPath_TUBEINSERTMach = ""
        gs_User_TUBEINSERTMach = ""
        gs_Password_TUBEINSERTMach = ""

    End Sub

    Private Sub up_TimeStart()
        m_Finish = False
        Me.Cursor = Cursors.WaitCursor

        Try
            ' Mixing Mach Vacuum HT42H
            '===========================
            Thread.Sleep(200)
            Thd_CAPCHECKERMac = New SchedulerSetting
            With Thd_CAPCHECKERMac
                .Name = "CAPCHECKERMac"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_CAPCHECKERMac
                .ScheduleThd = New Thread(AddressOf up_Refresh_CAPCHECKERMac)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "CAPCHECKERMac"
                .ScheduleThd.Start()
            End With

            ' Mixing Mach Vacuum HT35H
            '===========================
            Thread.Sleep(200)
            Thd_CELLSUPPLYBATTERYMach = New SchedulerSetting
            With Thd_CELLSUPPLYBATTERYMach
                .Name = "CELLSUPPLYBATTERYMach"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_CELLSUPPLYBATTERYMach
                .ScheduleThd = New Thread(AddressOf up_Refresh_CELLSUPPLYBATTERYMach)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "CELLSUPPLYBATTERYMach"
                .ScheduleThd.Start()
            End With

            ' Mixing Mach. Vacuum EC
            '===========================
            Thread.Sleep(200)
            Thd_RINGINSERTMach = New SchedulerSetting
            With Thd_RINGINSERTMach
                .Name = "RINGINSERTMach"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_RINGINSERTMach
                .ScheduleThd = New Thread(AddressOf up_Refresh_RINGINSERTMach)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "RINGINSERTMach"
                .ScheduleThd.Start()
            End With

            ' High Speed Mixer Mch
            '===========================
            Thread.Sleep(200)
            Thd_TUBEINSERTMach = New SchedulerSetting
            With Thd_TUBEINSERTMach
                .Name = "TUBEINSERTMach"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_TUBEINSERTMach
                .ScheduleThd = New Thread(AddressOf up_Refresh_TUBEINSERTMach)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "TUBEINSERTMach"
                .ScheduleThd.Start()
            End With

            ' Conveyor transfer powder
            '===========================
            Thread.Sleep(200)
            Thd_TUBESHRINKMach = New SchedulerSetting
            With Thd_TUBESHRINKMach
                .Name = "TUBESHRINKMach"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_TUBESHRINKMach
                .ScheduleThd = New Thread(AddressOf up_Refresh_TUBESHRINKMach)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "TUBESHRINKMach"
                .ScheduleThd.Start()
            End With

            ' Oscilator Mach
            '===========================
            Thread.Sleep(200)
            Thd_SIDECHECKERMach = New SchedulerSetting
            With Thd_SIDECHECKERMach
                .Name = "SIDECHECKERMach"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_SIDECHECKERMach
                .ScheduleThd = New Thread(AddressOf up_Refresh_SIDECHECKERMach)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "SIDECHECKERMach"
                .ScheduleThd.Start()
            End With

            ' Conveyor transfer box
            '===========================
            Thread.Sleep(200)
            Thd_CELLBOXINGMach = New SchedulerSetting
            With Thd_CELLBOXINGMach
                .Name = "CELLBOXINGMach"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_CELLBOXINGMach
                .ScheduleThd = New Thread(AddressOf up_Refresh_CELLBOXINGMach)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "CELLBOXINGMach"
                .ScheduleThd.Start()
            End With

            ' Box lift
            '===========================
            Thread.Sleep(200)
            Thd_SUPPLYBOXMACHMach = New SchedulerSetting
            With Thd_SUPPLYBOXMACHMach
                .Name = "SUPPLYBOXMACHMach"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_SUPPLYBOXMACHMach
                .ScheduleThd = New Thread(AddressOf up_Refresh_SUPPLYBOXMACHMach)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "SUPPLYBOXMACHMach"
                .ScheduleThd.Start()
            End With

            ' Expandmetal tension 1
            '===========================
            Thread.Sleep(200)
            Thd_CELLSUPPLYMach = New SchedulerSetting
            With Thd_CELLSUPPLYMach
                .Name = "CELLSUPPLYMach"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_CELLSUPPLYMach
                .ScheduleThd = New Thread(AddressOf up_Refresh_CELLSUPPLYMach)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "CELLSUPPLYMach"
                .ScheduleThd.Start()
            End With

            ' Expandmetal tension 2
            '===========================
            Thread.Sleep(200)
            Thd_UPPERRINGINSERTMach = New SchedulerSetting
            With Thd_UPPERRINGINSERTMach
                .Name = "UPPERRINGINSERTMach"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_UPPERRINGINSERTMach
                .ScheduleThd = New Thread(AddressOf up_Refresh_UPPERRINGINSERTMach)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "UPPERRINGINSERTMach"
                .ScheduleThd.Start()
            End With

            ' Oven Roller Mach. Manual #1
            '===========================
            Thread.Sleep(200)
            Thd_LABELLERMach = New SchedulerSetting
            With Thd_LABELLERMach
                .Name = "LABELLERMach"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_LABELLERMach
                .ScheduleThd = New Thread(AddressOf up_Refresh_LABELLERMach)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "LABELLERMach"
                .ScheduleThd.Start()
            End With

            ' Oven Heater Mach. Manual #1
            '===========================
            Thread.Sleep(200)
            Thd_OCVINSPECTIONMach = New SchedulerSetting
            With Thd_OCVINSPECTIONMach
                .Name = "OCVINSPECTIONMach"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_OCVINSPECTIONMach
                .ScheduleThd = New Thread(AddressOf up_Refresh_OCVINSPECTIONMach)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "OCVINSPECTIONMach"
                .ScheduleThd.Start()
            End With

            ' Oven Heater Mach. Manual #2
            '===========================
            Thread.Sleep(200)
            Thd_CELLBOXINGMach2 = New SchedulerSetting
            With Thd_CELLBOXINGMach2
                .Name = "CELLBOXINGMach2"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_Interval_CELLBOXINGMach2
                .ScheduleThd = New Thread(AddressOf up_Refresh_CELLBOXINGMach2)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "CELLBOXINGMach2"
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
        SyncLock Thd_CAPCHECKERMac.Lock
            Thd_CAPCHECKERMac.EndSchedule = True
        End SyncLock

        'Mixing Mach Vacuum HT35H
        SyncLock Thd_CELLSUPPLYBATTERYMach.Lock
            Thd_CELLSUPPLYBATTERYMach.EndSchedule = True
        End SyncLock

        'Mixing Mach Vacuum EC
        SyncLock Thd_RINGINSERTMach.Lock
            Thd_RINGINSERTMach.EndSchedule = True
        End SyncLock

        'High Speed Mixer Mch
        SyncLock Thd_TUBEINSERTMach.Lock
            Thd_TUBEINSERTMach.EndSchedule = True
        End SyncLock

        'Conveyor transfer powder
        SyncLock Thd_TUBESHRINKMach.Lock
            Thd_TUBESHRINKMach.EndSchedule = True
        End SyncLock

        'Oscilator Mach
        SyncLock Thd_SIDECHECKERMach.Lock
            Thd_SIDECHECKERMach.EndSchedule = True
        End SyncLock

        'Conveyor transfer box
        SyncLock Thd_CELLBOXINGMach.Lock
            Thd_CELLBOXINGMach.EndSchedule = True
        End SyncLock

        'Box lift
        SyncLock Thd_SUPPLYBOXMACHMach.Lock
            Thd_SUPPLYBOXMACHMach.EndSchedule = True
        End SyncLock

        'Expandmetal tension
        SyncLock Thd_CELLSUPPLYMach.Lock
            Thd_CELLSUPPLYMach.EndSchedule = True
        End SyncLock

        'UPPERRINGINSERTMach
        SyncLock Thd_UPPERRINGINSERTMach.Lock
            Thd_UPPERRINGINSERTMach.EndSchedule = True
        End SyncLock

        'OvenRoller Mach Manual 1
        SyncLock Thd_LABELLERMach.Lock
            Thd_LABELLERMach.EndSchedule = True
        End SyncLock

        'Oven Heater Mach Manual 1
        SyncLock Thd_OCVINSPECTIONMach.Lock
            Thd_OCVINSPECTIONMach.EndSchedule = True
        End SyncLock

        'OvenRoller Mach Manual 2
        SyncLock Thd_SIDEINSPECTIONMach.Lock
            Thd_SIDEINSPECTIONMach.EndSchedule = True
        End SyncLock

        'Oven Heater Mach Manual 2
        SyncLock Thd_CELLBOXINGMach2.Lock
            Thd_CELLBOXINGMach2.EndSchedule = True
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
    Private Sub up_Refresh_CAPCHECKERMac()
        up_Refresh(CAPCHECKERMac, "CAPCHECKERMac", Thd_CAPCHECKERMac, gs_ServerPath_CAPCHECKERMac, gs_LocalPath_CAPCHECKERMac, gs_User_CAPCHECKERMac, gs_Password_CAPCHECKERMac, gs_FileName_CAPCHECKERMac, Line_CAPCHECKERMac, Machine_CAPCHECKERMac)
    End Sub

    Private Sub up_Refresh_CELLSUPPLYBATTERYMach()
        up_Refresh(CELLSUPPLYBATTERYMach, "CELLSUPPLYBATTERYMach", Thd_CELLSUPPLYBATTERYMach, gs_ServerPath_CELLSUPPLYBATTERYMach, gs_LocalPath_CELLSUPPLYBATTERYMach, gs_User_CELLSUPPLYBATTERYMach, gs_Password_CELLSUPPLYBATTERYMach, gs_FileName_CELLSUPPLYBATTERYMach, Line_CELLSUPPLYBATTERYMach, Machine_CELLSUPPLYBATTERYMach)
    End Sub

    Private Sub up_Refresh_RINGINSERTMach()
        up_Refresh(RINGINSERTMach, "RINGINSERTMach", Thd_RINGINSERTMach, gs_ServerPath_RINGINSERTMach, gs_LocalPath_RINGINSERTMach, gs_User_RINGINSERTMach, gs_Password_RINGINSERTMach, gs_FileName_RINGINSERTMach, Line_RINGINSERTMach, Machine_RINGINSERTMach)
    End Sub

    Private Sub up_Refresh_TUBEINSERTMach()
        up_Refresh(TUBEINSERTMach, "TUBEINSERTMach", Thd_TUBEINSERTMach, gs_ServerPath_TUBEINSERTMach, gs_LocalPath_TUBEINSERTMach, gs_User_TUBEINSERTMach, gs_Password_TUBEINSERTMach, gs_FileName_TUBEINSERTMach, Line_TUBEINSERTMach, Machine_TUBEINSERTMach)
    End Sub

    Private Sub up_Refresh_TUBESHRINKMach()
        up_Refresh(TUBESHRINKMach, "TUBESHRINKMach", Thd_TUBESHRINKMach, gs_ServerPath_TUBESHRINKMach, gs_LocalPath_TUBESHRINKMach, gs_User_TUBESHRINKMach, gs_Password_TUBESHRINKMach, gs_FileName_TUBESHRINKMach, Line_TUBESHRINKMach, Machine_TUBESHRINKMach)
    End Sub

    Private Sub up_Refresh_SIDECHECKERMach()
        up_Refresh(SIDECHECKERMach, "SIDECHECKERMach", Thd_SIDECHECKERMach, gs_ServerPath_SIDECHECKERMach, gs_LocalPath_SIDECHECKERMach, gs_User_SIDECHECKERMach, gs_Password_SIDECHECKERMach, gs_FileName_SIDECHECKERMach, Line_SIDECHECKERMach, Machine_SIDECHECKERMach)
    End Sub

    Private Sub up_Refresh_CELLBOXINGMach()
        up_Refresh(CELLBOXINGMach, "CELLBOXINGMach", Thd_CELLBOXINGMach, gs_ServerPath_CELLBOXINGMach, gs_LocalPath_CELLBOXINGMach, gs_User_CELLBOXINGMach, gs_Password_CELLBOXINGMach, gs_FileName_CELLBOXINGMach, Line_CELLBOXINGMach, Machine_CELLBOXINGMach)
    End Sub

    Private Sub up_Refresh_SUPPLYBOXMACHMach()
        up_Refresh(SUPPLYBOXMACHMach, "SUPPLYBOXMACHMach", Thd_SUPPLYBOXMACHMach, gs_ServerPath_SUPPLYBOXMACHMach, gs_LocalPath_SUPPLYBOXMACHMach, gs_User_SUPPLYBOXMACHMach, gs_Password_SUPPLYBOXMACHMach, gs_FileName_SUPPLYBOXMACHMach, Line_SUPPLYBOXMACHMach, Machine_SUPPLYBOXMACHMach)
    End Sub

    Private Sub up_Refresh_CELLSUPPLYMach()
        up_Refresh(CELLSUPPLYMach, "CELLSUPPLYMach", Thd_CELLSUPPLYMach, gs_ServerPath_CELLSUPPLYMach, gs_LocalPath_CELLSUPPLYMach, gs_User_CELLSUPPLYMach, gs_Password_CELLSUPPLYMach, gs_FileName_CELLSUPPLYMach, Line_CELLSUPPLYMach, Machine_CELLSUPPLYMach)
    End Sub

    Private Sub up_Refresh_UPPERRINGINSERTMach()
        up_Refresh(UPPERRINGINSERTMach, "UPPERRINGINSERTMach", Thd_UPPERRINGINSERTMach, gs_ServerPath_UPPERRINGINSERTMach, gs_LocalPath_UPPERRINGINSERTMach, gs_User_UPPERRINGINSERTMach, gs_Password_UPPERRINGINSERTMach, gs_FileName_UPPERRINGINSERTMach, Line_UPPERRINGINSERTMach, Machine_UPPERRINGINSERTMach)
    End Sub

    Private Sub up_Refresh_LABELLERMach()
        up_Refresh(LABELLERMach, "LABELLERMach", Thd_LABELLERMach, gs_ServerPath_LABELLERMach, gs_LocalPath_LABELLERMach, gs_User_LABELLERMach, gs_Password_LABELLERMach, gs_FileName_LABELLERMach, Line_LABELLERMach, Machine_LABELLERMach)
    End Sub

    Private Sub up_Refresh_OCVINSPECTIONMach()
        up_Refresh(OCVINSPECTIONMach, "OCVINSPECTIONMach", Thd_OCVINSPECTIONMach, gs_ServerPath_OCVINSPECTIONMach, gs_LocalPath_OCVINSPECTIONMach, gs_User_OCVINSPECTIONMach, gs_Password_OCVINSPECTIONMach, gs_FileName_OCVINSPECTIONMach, Line_OCVINSPECTIONMach, Machine_OCVINSPECTIONMach)
    End Sub

    Private Sub up_Refresh_SIDEINSPECTIONMach()
        up_Refresh(SIDEINSPECTIONMach, "SIDEINSPECTIONMach", Thd_SIDEINSPECTIONMach, gs_ServerPath_SIDEINSPECTIONMach, gs_LocalPath_SIDEINSPECTIONMach, gs_User_SIDEINSPECTIONMach, gs_Password_SIDEINSPECTIONMach, gs_FileName_SIDEINSPECTIONMach, Line_SIDEINSPECTIONMach, Machine_SIDEINSPECTIONMach)
    End Sub

    Private Sub up_Refresh_CELLBOXINGMach2()
        up_Refresh(CELLBOXINGMach2, "CELLBOXINGMach2", Thd_CELLBOXINGMach2, gs_ServerPath_CELLBOXINGMach2, gs_LocalPath_CELLBOXINGMach2, gs_User_CELLBOXINGMach2, gs_Password_CELLBOXINGMach2, gs_FileName_CELLBOXINGMach2, Line_CELLBOXINGMach2, Machine_CELLBOXINGMach2)
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
        If gs_ServerPath_CAPCHECKERMac <> "" And gs_LocalPath_CAPCHECKERMac <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_CAPCHECKERMac, gs_User_CAPCHECKERMac, gs_Password_CAPCHECKERMac)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_CAPCHECKERMac, gs_Password_CAPCHECKERMac, gs_ServerPath_CAPCHECKERMac, gs_LocalPath_CAPCHECKERMac, gs_FileName_CAPCHECKERMac)
            End If
        End If
        '====================

        'Mixing Mach. Vacuum HT35H 
        '====================
        If gs_ServerPath_CELLSUPPLYBATTERYMach <> "" And gs_LocalPath_CELLSUPPLYBATTERYMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_CELLSUPPLYBATTERYMach, gs_User_CELLSUPPLYBATTERYMach, gs_Password_CELLSUPPLYBATTERYMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_CELLSUPPLYBATTERYMach, gs_Password_CELLSUPPLYBATTERYMach, gs_ServerPath_CELLSUPPLYBATTERYMach, gs_LocalPath_CELLSUPPLYBATTERYMach, gs_FileName_CELLSUPPLYBATTERYMach)
            End If
        End If
        '====================

        'Mixing Mach. Vacuum EC 
        '====================
        If gs_ServerPath_RINGINSERTMach <> "" And gs_LocalPath_RINGINSERTMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_RINGINSERTMach, gs_User_RINGINSERTMach, gs_Password_RINGINSERTMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_RINGINSERTMach, gs_Password_RINGINSERTMach, gs_ServerPath_RINGINSERTMach, gs_LocalPath_RINGINSERTMach, gs_FileName_RINGINSERTMach)
            End If
        End If
        '====================

        'High Speed Mixer Mch
        '====================
        If gs_ServerPath_TUBEINSERTMach <> "" And gs_LocalPath_TUBEINSERTMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_TUBEINSERTMach, gs_User_TUBEINSERTMach, gs_Password_TUBEINSERTMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_TUBEINSERTMach, gs_Password_TUBEINSERTMach, gs_ServerPath_TUBEINSERTMach, gs_LocalPath_TUBEINSERTMach, gs_FileName_TUBEINSERTMach)
            End If
        End If
        '====================

        'Conveyor transfer powder
        '====================
        If gs_ServerPath_TUBESHRINKMach <> "" And gs_LocalPath_TUBESHRINKMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_TUBESHRINKMach, gs_User_TUBESHRINKMach, gs_Password_TUBESHRINKMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_TUBESHRINKMach, gs_Password_TUBESHRINKMach, gs_ServerPath_TUBESHRINKMach, gs_LocalPath_TUBESHRINKMach, gs_FileName_TUBESHRINKMach)
            End If
        End If
        '====================

        'Oscilator Mach
        '====================
        If gs_ServerPath_SIDECHECKERMach <> "" And gs_LocalPath_SIDECHECKERMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_SIDECHECKERMach, gs_User_SIDECHECKERMach, gs_Password_SIDECHECKERMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_SIDECHECKERMach, gs_Password_SIDECHECKERMach, gs_ServerPath_SIDECHECKERMach, gs_LocalPath_SIDECHECKERMach, gs_FileName_SIDECHECKERMach)
            End If
        End If
        '====================

        'Conveyor transfer box
        '====================
        If gs_ServerPath_CELLBOXINGMach <> "" And gs_LocalPath_CELLBOXINGMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_CELLBOXINGMach, gs_User_CELLBOXINGMach, gs_Password_CELLBOXINGMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_CELLBOXINGMach, gs_Password_CELLBOXINGMach, gs_ServerPath_CELLBOXINGMach, gs_LocalPath_CELLBOXINGMach, gs_FileName_CELLBOXINGMach)
            End If
        End If
        '====================

        'Box lift
        '====================
        If gs_ServerPath_SUPPLYBOXMACHMach <> "" And gs_LocalPath_SUPPLYBOXMACHMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_SUPPLYBOXMACHMach, gs_User_SUPPLYBOXMACHMach, gs_Password_SUPPLYBOXMACHMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_SUPPLYBOXMACHMach, gs_Password_SUPPLYBOXMACHMach, gs_ServerPath_SUPPLYBOXMACHMach, gs_LocalPath_SUPPLYBOXMACHMach, gs_FileName_SUPPLYBOXMACHMach)
            End If
        End If
        '====================

        'Expandmetal tension 1
        '====================
        If gs_ServerPath_CELLSUPPLYMach <> "" And gs_LocalPath_CELLSUPPLYMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_CELLSUPPLYMach, gs_User_CELLSUPPLYMach, gs_Password_CELLSUPPLYMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_CELLSUPPLYMach, gs_Password_CELLSUPPLYMach, gs_ServerPath_CELLSUPPLYMach, gs_LocalPath_CELLSUPPLYMach, gs_FileName_CELLSUPPLYMach)
            End If
        End If
        '====================

        'UPPERRINGINSERTMach
        '====================
        If gs_ServerPath_UPPERRINGINSERTMach <> "" And gs_LocalPath_UPPERRINGINSERTMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_UPPERRINGINSERTMach, gs_User_UPPERRINGINSERTMach, gs_Password_UPPERRINGINSERTMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_UPPERRINGINSERTMach, gs_Password_UPPERRINGINSERTMach, gs_ServerPath_UPPERRINGINSERTMach, gs_LocalPath_UPPERRINGINSERTMach, gs_FileName_UPPERRINGINSERTMach)
            End If
        End If
        '====================

        'Oven Roller Mach. Manual #1
        '====================
        If gs_ServerPath_LABELLERMach <> "" And gs_LocalPath_LABELLERMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_LABELLERMach, gs_User_LABELLERMach, gs_Password_LABELLERMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_LABELLERMach, gs_Password_LABELLERMach, gs_ServerPath_LABELLERMach, gs_LocalPath_LABELLERMach, gs_FileName_LABELLERMach)
            End If
        End If
        '====================

        'Oven Heater Mach. Manual #1
        '====================
        If gs_ServerPath_OCVINSPECTIONMach <> "" And gs_LocalPath_OCVINSPECTIONMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_OCVINSPECTIONMach, gs_User_OCVINSPECTIONMach, gs_Password_OCVINSPECTIONMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_OCVINSPECTIONMach, gs_Password_OCVINSPECTIONMach, gs_ServerPath_OCVINSPECTIONMach, gs_LocalPath_OCVINSPECTIONMach, gs_FileName_OCVINSPECTIONMach)
            End If
        End If
        '====================

        'Oven Roller Mach. Manual #2
        '====================
        If gs_ServerPath_SIDEINSPECTIONMach <> "" And gs_LocalPath_SIDEINSPECTIONMach <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_SIDEINSPECTIONMach, gs_User_SIDEINSPECTIONMach, gs_Password_SIDEINSPECTIONMach)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_SIDEINSPECTIONMach, gs_Password_SIDEINSPECTIONMach, gs_ServerPath_SIDEINSPECTIONMach, gs_LocalPath_SIDEINSPECTIONMach, gs_FileName_SIDEINSPECTIONMach)
            End If
        End If
        '====================

        'Oven Heater Mach. Manual #2
        '====================
        If gs_ServerPath_CELLBOXINGMach2 <> "" And gs_LocalPath_CELLBOXINGMach2 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPath_CELLBOXINGMach2, gs_User_CELLBOXINGMach2, gs_Password_CELLBOXINGMach2)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_User_CELLBOXINGMach2, gs_Password_CELLBOXINGMach2, gs_ServerPath_CELLBOXINGMach2, gs_LocalPath_CELLBOXINGMach2, gs_FileName_CELLBOXINGMach2)
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