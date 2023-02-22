Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Threading
Imports C1.Win.C1FlexGrid

Public Class FormSchedulerCSV_CaseSealant_Trouble
#Region "INITIAL"

    Private cConfig As clsConfig

    Const ConnectionErrorMsg As String = "A network-related or instance-specific error occurred while establishing a connection to SQL Server"
    Const TransportErrorMsg As String = "A transport-level error has occurred"

    Public Sub New()
        InitializeComponent()
    End Sub

#End Region

#Region "DECLARATION"

    Dim CaseSealant1 As Byte = 1
    Dim CaseSealant2 As Byte = 2


    Dim col_ProcessName As Integer = 0
    Dim Col_ProcessType As Byte = 1
    Dim Col_ProcessStatus As Byte = 2
    Dim Col_LastProcess As Byte = 3
    Dim Col_NextProcess As Byte = 4
    Dim Col_ErrorMessage As Byte = 5

    Dim col_Count As Integer = 6

    Dim FilesList As New List(Of String)

    Dim Thd_Mix01 As SchedulerSetting
    Dim Thd_Mix02 As SchedulerSetting


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

        Do Until Thd_Mix01.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_Mix01.Status = "iddle" Then
                Thd_Mix01.ScheduleThd = Nothing
            End If
            If Thd_Mix01.ScheduleThd Is Nothing Then Exit Do
            Thread.Sleep(100)
        Loop

        Do Until Thd_Mix02.ScheduleThd.ThreadState = Threading.ThreadState.Stopped
            If Thd_Mix02.Status = "iddle" Then
                Thd_Mix02.ScheduleThd = Nothing
            End If
            If Thd_Mix02.ScheduleThd Is Nothing Then Exit Do
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

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        ShowInTaskbar = True
        Me.WindowState = FormWindowState.Normal
        NotifyIcon1.Visible = False
    End Sub

    Private Sub timerCurr_Tick(sender As Object, e As EventArgs) Handles timerCurr.Tick
        lblCurrTime.Text = Format(Now, "HH:mm:ss")
        lblCurrDate.Text = Format(Now, "dddd , dd MMM yyyy")
    End Sub

    Private Sub FormSchedulerCSV_CaseSealant_Trouble_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

    Private Sub FormSchedulerCSV_CaseSealant_Trouble_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                'NotifyIcon1.Icon = SystemIcons.Application
                NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
                NotifyIcon1.BalloonTipTitle = "RELAY PLC CSV TROUBLE (CaseSealant)"
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

            '01. CaseSealant X-01
            pProses = "Holder insert S-4"
            up_ProcessData(CaseSealant1, "Holder insert S-4", gs_LocalPathMix01, gs_FileName_TroubleMix01, "", 1)

            '02. CaseSealant X-02
            pProses = "Sealant coating S-4"
            up_ProcessData(CaseSealant2, "Sealant coating S-4", gs_LocalPathMix02, gs_FileName_TroubleMix02, "", 1)


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

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "Holder insert S-4" Then
                            gs_ServerPathMix01 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathMix01 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserMix01 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordMix01 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalMix01 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If

                        If ds.Tables(0).Rows(i)("Machine_Type").ToString.Trim = "Sealant coating S-4" Then
                            gs_ServerPathMix02 = ds.Tables(0).Rows(i)("Server_Path").ToString.Trim
                            gs_LocalPathMix02 = ds.Tables(0).Rows(i)("Local_Path").ToString.Trim
                            gs_UserMix02 = ds.Tables(0).Rows(i)("user_FTP").ToString.Trim
                            gs_PasswordMix02 = ds.Tables(0).Rows(i)("Password_FTP").ToString.Trim
                            gi_IntervalMix02 = ds.Tables(0).Rows(i)("TimeInterval")
                        End If


                    Next

                End If

            End Using

        Catch ex As Exception
            WriteToErrorLog("Load data FTP error", ex.Message)
        End Try

    End Sub

    Private Sub up_ClearVariable()

        gs_ServerPathMix01 = ""
        gs_LocalPathMix01 = ""
        gs_UserMix01 = ""
        gs_PasswordMix01 = ""

        gs_ServerPathMix02 = ""
        gs_LocalPathMix02 = ""
        gs_UserMix02 = ""
        gs_PasswordMix02 = ""


    End Sub

#Region "PROCESS"

    Private Sub up_TimeStart()
        m_Finish = False
        Me.Cursor = Cursors.WaitCursor

        Try

            Thread.Sleep(200)
            Thd_Mix01 = New SchedulerSetting
            With Thd_Mix01
                .Name = "Mix01"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalMix01
                .ScheduleThd = New Thread(AddressOf up_RefreshCaseSealant01)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "Mix01"
                .ScheduleThd.Start()
            End With

            Thread.Sleep(220)
            Thd_Mix02 = New SchedulerSetting
            With Thd_Mix02
                .Name = "Mix02"
                .EndSchedule = False
                .Lock = New Object
                .DelayTime = gi_IntervalMix02
                .ScheduleThd = New Thread(AddressOf up_RefreshCaseSealant02)
                .ScheduleThd.IsBackground = True
                .ScheduleThd.Name = "Mix02"
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

        SyncLock Thd_Mix01.Lock
            Thd_Mix01.EndSchedule = True
        End SyncLock

        SyncLock Thd_Mix02.Lock
            Thd_Mix02.EndSchedule = True
        End SyncLock


    End Sub

    Private Function up_ToInserDatatable_Trouble(ByVal pLocalPath As String, ByVal pFileName As String, ByVal pLineCode As String, ByVal pGroupCount As Integer) As DataTable
        Dim dt As New DataTable
        Dim tmpDatde As String()
        Dim Col_Line As String = "", AlarmCode As String
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
                            'dt.Rows.Add(
                            '    Trim(dt_LastData.Rows(0).Item("LineCode")),
                            '    Trim(dt_LastData.Rows(0).Item("Mode")),
                            '    Trim(dt_LastData.Rows(0).Item("Status")),
                            '    Trim(dt_LastData.Rows(0).Item("TroubleCode")),
                            '    Trim(dt_LastData.Rows(0).Item("Start_Time")),
                            '    Trim(dt_LastData.Rows(0).Item("End_Time")),
                            '    Trim(dt_LastData.Rows(0).Item("Last_Update")),
                            '    Trim(dt_LastData.Rows(0).Item("Last_User"))
                            '    )

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

                                Dim group As Int16
                                group = GetGroup(dtCSV, x, pGroupTroubleCount)

                                EndTime = ""

                                Dim startGroup As Int16, endGroup As Int16

                                If (group <> 0) Then

                                    startGroup = Convert.ToInt16(list(group - 1).Split(":")(0))
                                    endGroup = Convert.ToInt16(list(group - 1).Split(":")(1))

                                    If x = 0 Then 'Row Pertama otomatis loging
                                        AlarmCode = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(group + 3))), startGroup, endGroup)
                                        dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                        iRow_Data = iRow_Data + 1
                                    Else

                                        If (tmp_Alarm3 <> ls_Alarm3 And ls_Trouble = False) Then
                                            EndTime = StartTime
                                            dt.Rows(iRow_Data - 1)("EndTime") = StartTime
                                            AlarmCode = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(group + 3))), startGroup, endGroup)
                                            dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                            iRow_Data = iRow_Data + 1

                                            'Dim pLastBit As Integer = up_GetLastBit(AlarmCode, startGroup, endGroup)
                                            'Dim pMid As Integer = 16
                                            'Dim pChr As String = ""

                                            'For i As Integer = startGroup To endGroup
                                            '    pChr = Mid(AlarmCode, pMid, 1)

                                            '    If pChr <> "0" Then

                                            '        AlarmCode = i

                                            '        If i <> pLastBit Then
                                            '            EndTime = StartTime
                                            '        Else
                                            '            EndTime = ""
                                            '        End If

                                            '        dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                            '        iRow_Data = iRow_Data + 1

                                            '    End If
                                            '    pMid = pMid - 1
                                            'Next

                                            ls_Trouble = True

                                        End If

                                    End If
                                Else
                                    EndTime = ""

                                    If x = 0 Then 'Row Pertama otomatis loging
                                        startGroup = Convert.ToInt16(list(pGroupTroubleCount - 1).Split(":")(0))
                                        endGroup = Convert.ToInt16(list(pGroupTroubleCount - 1).Split(":")(1))

                                        AlarmCode = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(pGroupTroubleCount + 3))), startGroup, endGroup)
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

                                'JIKA GROUP 3 ADA ALARM, MAKA AMBIL LEBIH DULU'
                                '=============================================
                                'If ls_Alarm3 <> 0 Then

                                '    EndTime = ""

                                '    'JIKA GROUP 3 ADA ALARM, MAKA AMBIL LEBIH DULU'
                                '    '=============================================
                                '    If x = 0 Then 'Row Pertama otomatis loging

                                '        AlarmCode = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(6))), 32, 47)

                                '        'If (tmp_Alarm3 <> ls_Alarm3 And ls_Alarm3 = 0) Then
                                '        '    dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '        '    iRow_Data = iRow_Data + 1
                                '        'ElseIf (tmp_Alarm3 = "0" And ls_Alarm3 <> 0) Then
                                '        '    dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '        '    iRow_Data = iRow_Data + 1
                                '        'End If

                                '        dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '        iRow_Data = iRow_Data + 1
                                '    Else

                                '        If (tmp_Alarm3 <> ls_Alarm3 And ls_Trouble = False) Then

                                '            dt.Rows(iRow_Data - 1)("EndTime") = StartTime

                                '            Dim pLastBit As Integer = up_GetLastBit(ls_Alarm3, 32, 47)
                                '            Dim pMid As Integer = 16
                                '            Dim pChr As String = ""

                                '            For i As Integer = 32 To 47
                                '                pChr = Mid(ls_Alarm3, pMid, 1)

                                '                If pChr <> "0" Then

                                '                    AlarmCode = i

                                '                    If i <> pLastBit Then
                                '                        EndTime = StartTime
                                '                    Else
                                '                        EndTime = ""
                                '                    End If

                                '                    dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '                    iRow_Data = iRow_Data + 1

                                '                End If
                                '                pMid = pMid - 1
                                '            Next

                                '            ls_Trouble = True

                                '        End If

                                '    End If
                                '    'End If

                                'ElseIf ls_Alarm2 <> 0 Then

                                '    EndTime = ""

                                '    'NEXT JIKA ALARM GROUP 3 TIDAK ADA DAN GROUP 2 ADA ALARM MAKA AMBIL GROUP 2
                                '    '========================================================================
                                '    If x = 0 Then 'Row Pertama otomatis loging

                                '        AlarmCode = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(5))), 16, 31)

                                '        'If (tmp_Alarm2 <> AlarmCode And AlarmCode = 0) Then
                                '        '    dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '        '    iRow_Data = iRow_Data + 1
                                '        'ElseIf (tmp_Alarm2 = "0" And AlarmCode <> 0) Then
                                '        '    dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '        '    iRow_Data = iRow_Data + 1

                                '        'End If

                                '        dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '        iRow_Data = iRow_Data + 1
                                '    Else

                                '        If (tmp_Alarm2 <> ls_Alarm2 And ls_Trouble = False) Then

                                '            dt.Rows(iRow_Data - 1)("EndTime") = StartTime

                                '            Dim pLastBit As Integer = up_GetLastBit(ls_Alarm2, 16, 31)
                                '            Dim pMid As Integer = 16
                                '            Dim pChr As String = ""

                                '            For i As Integer = 16 To 31
                                '                pChr = Mid(ls_Alarm2, pMid, 1)

                                '                If pChr <> "0" Then

                                '                    AlarmCode = i

                                '                    If i <> pLastBit Then
                                '                        EndTime = StartTime
                                '                    Else
                                '                        EndTime = ""
                                '                    End If

                                '                    dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '                    iRow_Data = iRow_Data + 1

                                '                End If
                                '                pMid = pMid - 1
                                '            Next

                                '            ls_Trouble = True

                                '        End If

                                '    End If
                                '    'End If

                                'ElseIf ls_Alarm1 <> 0 Then

                                '    EndTime = ""

                                '    'NEXT JIKA GROUP 3 DAN 2 TIDAK ADA ALARM MAKA AMBIL GROUP 1
                                '    '==========================================================
                                '    If x = 0 Then 'Row Pertama otomatis loging

                                '        AlarmCode = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(4))), 0, 15)

                                '        'If (tmp_Alarm1 <> ls_Alarm1 And ls_Alarm1 = 0) Then
                                '        '    dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '        '    iRow_Data = iRow_Data + 1
                                '        'ElseIf (tmp_Alarm1 = "0" And ls_Alarm1 <> 0) Then
                                '        '    dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '        '    iRow_Data = iRow_Data + 1
                                '        'End If

                                '        dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '        iRow_Data = iRow_Data + 1
                                '    Else

                                '        If (tmp_Alarm1 <> ls_Alarm1 And ls_Trouble = False) Then

                                '            dt.Rows(iRow_Data - 1)("EndTime") = StartTime

                                '            Dim pLastBit As Integer = up_GetLastBit(ls_Alarm1, 0, 15)
                                '            Dim pMid As Integer = 16
                                '            Dim pChr As String = ""

                                '            For i As Integer = 0 To 15
                                '                pChr = Mid(ls_Alarm1, pMid, 1)

                                '                If pChr <> "0" Then

                                '                    AlarmCode = i

                                '                    If i <> pLastBit Then
                                '                        EndTime = StartTime
                                '                    Else
                                '                        EndTime = ""
                                '                    End If

                                '                    dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '                    iRow_Data = iRow_Data + 1

                                '                End If
                                '                pMid = pMid - 1
                                '            Next

                                '            ls_Trouble = True

                                '        End If

                                '    End If

                                '    'End If


                                'ElseIf ls_Alarm1 = 0 And ls_Alarm2 = 0 And ls_Alarm3 = 0 Then



                                '    EndTime = ""

                                '    If x = 0 Then 'Row Pertama otomatis loging

                                '        AlarmCode = up_GetCodeTrouble(DecimalToBinary(Trim(dtCSV.Rows(x)(6))), 32, 47)

                                '        'If tmp_Alarm1 <> "0" Then
                                '        '    dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '        '    iRow_Data = iRow_Data + 1
                                '        'End If

                                '        dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '        iRow_Data = iRow_Data + 1
                                '    Else

                                '        If tmp_Alarm1 <> ls_Alarm1 Or tmp_Alarm2 <> ls_Alarm2 Or tmp_Alarm3 <> ls_Alarm3 Then

                                '            dt.Rows(iRow_Data - 1)("EndTime") = StartTime

                                '            AlarmCode = 0

                                '            dt.Rows.Add(Col_Line, ModeCls, StatusCls, AlarmCode, StartTime, EndTime, LastUpdate, "Tes")
                                '            iRow_Data = iRow_Data + 1

                                '            ls_Trouble = False

                                '        End If

                                '    End If



                                'End If

                                tmp_StartTime = StartTime
                                tmp_Alarm1 = ls_Alarm1
                                tmp_Alarm2 = ls_Alarm2
                                tmp_Alarm3 = AlarmCode

                                'End If

                            End If
                        Next

                        'Send to History Folder
                        '======================
                        'My.Computer.FileSystem.MoveFile(pLocalPath & "\" & fi.Name, pLocalPath & "\History\" & fi.Name, True)

                        Return dt

                    End If

                End If

            Next

        Catch ex As Exception
            WriteToErrorLog("Trouble Process 2 error in LineCode : " & pLineCode, ex.Message)
        End Try

    End Function

    Private Sub up_RefreshCaseSealant01()
        up_Refresh(CaseSealant1, "Holder insert S-4", Thd_Mix01, gs_ServerPathMix01, gs_LocalPathMix01, gs_UserMix01, gs_PasswordMix01, gs_FileName_TroubleMix01, 1, "", "")
    End Sub
    Private Sub up_RefreshCaseSealant02()
        up_Refresh(CaseSealant2, "Sealant coating S-4", Thd_Mix02, gs_ServerPathMix02, gs_LocalPathMix02, gs_UserMix02, gs_PasswordMix02, gs_FileName_TroubleMix02, 1, "", "")
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

        grid.Item(pProcess, Col_ErrorMessage) = ""

        Dim dtHis, dtHis2 As New DataTable
        dtHis = up_ToInserDatatable_Trouble(pLocalPath, pFileNameHistory, pLineCode, pGroupCount)
        dtHis2 = up_ToInserDatatable_Trouble2(pLocalPath, pFileNameHistory, pLineCode, pGroupCount)

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


    Private Sub up_FTP()

        'FTP CaseSealant 1
        '====================
        If gs_ServerPathMix01 <> "" And gs_LocalPathMix01 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathMix01, gs_UserMix01, gs_PasswordMix01)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserMix01, gs_PasswordMix01, gs_ServerPathMix01, gs_LocalPathMix01, gs_FileName_TroubleMix01, "zz")
            End If
        End If
        '====================

        'FTP CaseSealant 2
        '====================
        If gs_ServerPathMix02 <> "" And gs_LocalPathMix02 <> "" Then
            FilesList = GetFtpFileList(gs_ServerPathMix02, gs_UserMix02, gs_PasswordMix02)
            If FilesList.Count > 0 Then
                DownloadFiles(FilesList, "", gs_UserMix02, gs_PasswordMix02, gs_ServerPathMix02, gs_LocalPathMix02, gs_FileName_TroubleMix02, "zz")
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