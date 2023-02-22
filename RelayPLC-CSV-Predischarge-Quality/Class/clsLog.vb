Imports System.IO
Imports System.Text
Imports System.Diagnostics
Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class clsLog

#Region "Declaration"

    Private UserLogin As String
    Private ConStr As String
    Private li_StartTime As Date
    Private li_EndTime As Date
    Private li_Duration As TimeSpan
    Private dtMsg As DataTable
    Const ConnectionErrorMsg As String = "A network-related or instance-specific error occurred while establishing a connection to SQL Server"

    Public Enum ErrSeverity
        ALERT = 1
        ERR = 2
        INFO = 3
    End Enum

#End Region

#Region "Initialization"

    Public Sub New(ByVal pConStr As String, ByVal pUserlogin As String)
        ' Add any initialization after the InitializeComponent() call.
        ConStr = pConStr
        UserLogin = pUserlogin
        up_GetMsg()
    End Sub

#End Region

#Region "Procedures and Functions"

    ''' <summary>
    ''' Get all data from Message table and store it in a datatable.
    ''' This function is called once when the class is initialized.
    ''' Later this datatable is accessed to get messages without having to connect to database.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub up_GetMsg()
        Try
            dtMsg = New DataTable
            Using Cn As New SqlConnection(ConStr)
                Cn.Open()
                Dim cmd As New SqlCommand("select * from ErrorLog", Cn)
                Dim da As New SqlDataAdapter(cmd)
                da.Fill(dtMsg)
                da.Dispose()
                da = Nothing
            End Using
        Catch ex As Exception
            If ex.Message.Contains(ConnectionErrorMsg) Then
                Throw New Exception("Database connection failed")
            Else
                Throw ex
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Gets message description from database based on message ID.
    ''' </summary>
    ''' <param name="pLogID">
    ''' Log ID.
    ''' </param>
    ''' <returns>
    ''' Returns message description.
    ''' </returns>
    ''' <remarks></remarks>
    Public Function uf_ConvertMsg(ByVal pLogID As String) As String
        Try
            Dim Msg As String = ""
            Dim column(0) As DataColumn
            column(0) = dtMsg.Columns(0)
            dtMsg.PrimaryKey = column

            Dim row As DataRow = dtMsg.Rows.Find(pLogID)
            If row IsNot Nothing AndAlso row(1) IsNot Nothing Then
                Msg = row(1).ToString.TrimEnd
            End If
            Return Msg
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    ''' <summary>
    ''' Gets event ID from database based on Log ID.
    ''' </summary>
    ''' <param name="pLogID">
    ''' Log ID.
    ''' </param>
    ''' <returns>
    ''' Returns Event ID
    ''' </returns>
    ''' <remarks></remarks>
    Private Function uf_ConvertEventID(ByVal pLogID As String) As Integer
        Try
            Dim pEventID As Integer
            pEventID = 0
            Dim column(0) As DataColumn
            column(0) = dtMsg.Columns(0)
            dtMsg.PrimaryKey = column

            Dim row As DataRow = dtMsg.Rows.Find(pLogID)
            If row IsNot Nothing AndAlso row(3) IsNot Nothing Then
                pEventID = Val(row(3).ToString)
            End If
            Return pEventID
        Catch ex As Exception
            Return ex.Message
        End Try

    End Function

    '*************************************************************
    'NAME:          WriteToErrorLog
    'PURPOSE:       Open or create an error log and submit error message
    'PARAMETERS:    pScreenName     - Name of Screen Process
    '               pErrSummary     - Err Summary Process
    '               pErrID          - Error ID
    '               pErrSeverity    - Enum ALERT, INFO and Err of the error file entry
    'RETURNS:       Nothing
    '*************************************************************
    ''' <summary>
    ''' this procedure for process error log
    ''' </summary>
    ''' <param name="pScreenName">Screen Name</param>
    ''' <param name="pErrSummary">Err Summary</param>
    ''' <param name="pErrID">Err ID</param>
    ''' <param name="pErrSeverity">Err Severity</param>
    ''' <remarks>
    ''' 1. create \Log\ directory if not create before
    ''' 2. create \Log\Process\ directory if not create before
    ''' 3. checking the process log file with the same date if found it will be add the process information with the exist file
    ''' 4. create the process detail information
    ''' 5. keep the process log file from last 90 days until now otherwise will removed
    ''' </remarks>
    Public Sub WriteToErrorLog(ByVal pScreenName As String, Optional ByVal pErrSummary As String = "", Optional ByVal pErrID As Integer = 9999, Optional ByVal pErrSeverity As ErrSeverity = ErrSeverity.ERR, Optional ByVal pHarigami As Boolean = False)

        Dim ls_Date As String
        Dim ls_ErrType As String
        Dim ls_dateFolder As String
        Dim ls_CompName As String = uf_CompName()
        Dim ls_LogFolder As String = "D:\RelayPLC"
        ls_dateFolder = Format(Now, "yyyyMMdd")
        ls_Date = Format(Now, "yyyyMMdd")
        ls_ErrType = uf_ErrSeverity(pErrSeverity)
        pScreenName = pScreenName.Trim
        pScreenName = pScreenName.Replace(" ", "_")


        'If Not System.IO.Directory.Exists("D:\") Then
        ls_LogFolder = Application.StartupPath
        'End If

        If pHarigami Then
            If Not System.IO.Directory.Exists(ls_LogFolder) Then
                System.IO.Directory.CreateDirectory(ls_LogFolder)
            End If

            If Not System.IO.Directory.Exists(ls_LogFolder & "\Log") Then
                System.IO.Directory.CreateDirectory(ls_LogFolder & "\Log")
            End If

            If Not System.IO.Directory.Exists(ls_LogFolder &
            "\Log\" & ls_dateFolder & "\Error") Then
                System.IO.Directory.CreateDirectory(ls_LogFolder &
                "\Log\" & ls_dateFolder & "\Error")
            End If

            If Not System.IO.Directory.Exists(ls_LogFolder &
            "\Log\" & ls_dateFolder & "\Error") Then
                System.IO.Directory.CreateDirectory(ls_LogFolder &
                "\Log\" & ls_dateFolder & "\Error")
            End If

            'check the file
            Dim fs As FileStream = New FileStream(ls_LogFolder &
            "\Log\" & ls_dateFolder & "\Error\err" & pScreenName & "_" & ls_Date & ".log", FileMode.OpenOrCreate, FileAccess.ReadWrite)
            Dim s As StreamWriter = New StreamWriter(fs)

            s.Close()
            fs.Close()

            'log it
            Dim fs1 As FileStream = New FileStream(ls_LogFolder &
            "\Log\" & ls_dateFolder & "\Error\err" & pScreenName & "_" & ls_Date & ".log", FileMode.Append, FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)

            s1.Write("" & Format(Now, "dd/MM/yyyy HH:mm:ss") & " ")
            s1.Write("[" & UserLogin & "]" & " ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & "" & " ")
            s1.Write("[" & ls_ErrType & "]" & " ")
            s1.Write("" & pErrSummary & "" & "")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()

        Else
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
            s1.Write("[" & UserLogin & "]" & " ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & "" & " ")
            s1.Write("[" & ls_ErrType & "]" & " ")
            s1.Write("" & pErrSummary & "" & "")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        End If

        'write to event log
        If ls_ErrType = "E" Then
            WriteToEventLog(pErrSummary, "RelayPLC(" & pScreenName & ")", EventLogEntryType.Error, "Application", pErrID)
        ElseIf ls_ErrType = "A" Then
            WriteToEventLog(pErrSummary, "RelayPLC(" & pScreenName & ")", EventLogEntryType.Warning, "Application", pErrID)
        End If

        'Dim ls_Dir As New IO.DirectoryInfo(ls_LogFolder & "\Log\Error\")
        'Dim ls_GetFile As IO.FileInfo() = ls_Dir.GetFiles()
        'Dim ls_File As IO.FileInfo
        'Dim li_index As Long
        'Dim ls_log As String
        'Dim li_CountDate As Long
        'Dim ls_Temp As List(Of String) = New List(Of String)

        'li_index = 0
        'For Each ls_File In ls_GetFile

        '    ls_Temp.Add(ls_File.ToString)

        '    ls_log = Right(ls_Temp.Item(li_index), 12)
        '    ls_log = Mid(ls_log, 1, 8)
        '    ls_log = Mid(ls_log, 5, 2) & "/" & Right(ls_log, 2) & "/" & Left(ls_log, 4)
        '    li_CountDate = DateDiff(DateInterval.Day, CDate(Format(CDate(ls_log), "MM/dd/yyyy")), CDate(Format(Now, "MM/dd/yyyy")))
        '    If li_CountDate > 90 Then
        '        File.Delete(ls_LogFolder & "\Log\Error\err" & pScreenName & "_" & Format(CDate(ls_log), "yyyyMMdd") & ".log")
        '    End If

        '    li_index = li_index + 1
        'Next
    End Sub


    '*************************************************************
    'NAME:          WriteToProcessLog
    'PURPOSE:       Open or create an error log and submit error message
    'PARAMETERS:    pStartTime      - Start of time process of the error file entry
    '               pScreenName     - Name of Screen Process
    '               pCustomMsg      - Name of Process of the error file entry
    '               pErrSummary     - Err summary process
    '               pErrID          - Error ID
    '               pErrSeverity    - Enum ALERT, INFO and Err of the error file entry
    '               pWriteEventLog  - True false for Write to Event Log
    '               pStartEndStatus - Different Format for start and end
    '               pUseLogTime     - True or False use log time
    'RETURNS:       Nothing
    '*************************************************************
    ''' <summary>
    ''' this procedure for process error log
    ''' </summary>
    ''' <param name="pStartTime">Start Time</param>
    ''' <param name="pScreenName">Screen Name</param>
    ''' <param name="pCustomMsg">Message</param>
    ''' <param name="pErrSummary">Err Summary</param>
    ''' <param name="pErrID">Err ID</param>
    ''' <param name="pErrSeverity">Err Severity</param>
    ''' <param name="pWriteToEventLog">Write to Event Log</param>
    ''' <param name="pStartEndStatus">Start or End Status</param>
    ''' <param name="pUseLogTime">Use Log Time</param>
    ''' <remarks>
    ''' 1. create \Log\ directory if not create before
    ''' 2. create \Log\Process\ directory if not create before
    ''' 3. checking the process log file with the same date if found it will be add the process information with the exist file
    ''' 4. create the process detail information
    ''' 5. keep the process log file from last 90 days until now otherwise will removed
    ''' </remarks>
    Public Sub WriteToProcessLog(ByVal pStartTime As Date,
                                 ByVal pScreenName As String,
                                 Optional ByVal pCustomMsg As String = "",
                                 Optional ByVal pErrSummary As String = "",
                                 Optional ByVal pErrID As Integer = 9999,
                                 Optional ByVal pErrSeverity As ErrSeverity = ErrSeverity.INFO,
                                 Optional ByVal pWriteToEventLog As Boolean = False,
                                 Optional ByVal pStartEndStatus As String = "",
                                 Optional ByVal pUseLogTime As Boolean = False)

        Dim ls_Date As String
        Dim ls_ErrType As String
        Dim ls_Duration As String
        Dim ls_CompName As String = uf_CompName()
        Dim ls_LogFolder As String = "D:\RelayPLC"


        'If Not System.IO.Directory.Exists("D:\") Then
        ls_LogFolder = Application.StartupPath
        'End If

        li_StartTime = pStartTime
        li_EndTime = Now
        li_Duration = li_EndTime - li_StartTime
        ls_Duration = uf_AddSpace(Format(li_Duration.TotalMilliseconds, "###.#0") & " (ms)", 15)

        ls_Date = Format(Now, "yyyyMMdd")
        ls_ErrType = uf_ErrSeverity(pErrSeverity)
        pScreenName = pScreenName.Trim
        pScreenName = pScreenName.Replace(" ", "_")

        If Not System.IO.Directory.Exists(ls_LogFolder) Then
            System.IO.Directory.CreateDirectory(ls_LogFolder)
        End If

        If Not System.IO.Directory.Exists(ls_LogFolder & "\Log") Then
            System.IO.Directory.CreateDirectory(ls_LogFolder & "\Log")
        End If

        If Not System.IO.Directory.Exists(ls_LogFolder & "\Log" &
        "\Process") Then
            System.IO.Directory.CreateDirectory(ls_LogFolder & "\Log" &
            "\Process")
        End If

        'check the file
        Dim fs As FileStream = New FileStream(ls_LogFolder &
        "\Log\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim s As StreamWriter = New StreamWriter(fs)

        s.Close()
        fs.Close()

        'log it
        Dim fs1 As FileStream = New FileStream(ls_LogFolder &
        "\Log\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)

        If pStartEndStatus.ToUpper.Trim = "START" Then
            If pUseLogTime = True Then
                s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
            Else
                s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
            End If
            s1.Write("[" & UserLogin & "] ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & " ")
            s1.Write("[" & ls_ErrType & "] ")
            s1.Write("start.")
            s1.Write("" & vbCrLf)
            s1.Write("* * * * *")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        ElseIf pStartEndStatus.ToUpper.Trim = "END" Then
            s1.Write("* * * * *")
            s1.Write("" & vbCrLf)
            If pUseLogTime = True Then
                s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
            Else
                s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
            End If
            s1.Write("[" & UserLogin & "] ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & " ")
            s1.Write("[" & ls_ErrType & "] ")
            s1.Write("end.")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        ElseIf pStartEndStatus = "" Then
            If ls_ErrType = "I" Then
                If pUseLogTime = True Then
                    s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
                Else
                    s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
                End If
                s1.Write("[" & UserLogin & "] ")
                s1.Write("[" & ls_CompName & "] ")
                s1.Write("" & pScreenName & " ")
                s1.Write("[" & ls_ErrType & "] ")
                s1.Write("" & pCustomMsg & " ")
                s1.Write("" & pErrSummary & " ")
                s1.Write("" & ls_Duration & "")
                s1.Write("" & vbCrLf)
                s1.Close()
                fs1.Close()
            Else
                If pUseLogTime = True Then
                    s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
                Else
                    s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
                End If
                s1.Write("[" & UserLogin & "] ")
                s1.Write("[" & ls_CompName & "] ")
                s1.Write("" & pScreenName & " ")
                s1.Write("[" & ls_ErrType & "] ")
                s1.Write("" & pCustomMsg & " ")
                s1.Write("" & pErrSummary & "")
                s1.Write("" & vbCrLf)
                s1.Close()
                fs1.Close()
            End If
        End If

        'write to event log
        If pWriteToEventLog = True Then
            If ls_ErrType = "E" Then
                WriteToEventLog(pCustomMsg, "RelayPLC(" & pScreenName & ")", EventLogEntryType.Error, "Application", pErrID)
            ElseIf ls_ErrType = "A" Then
                WriteToEventLog(pCustomMsg, "RelayPLC(" & pScreenName & ")", EventLogEntryType.Warning, "Application", pErrID)
            End If
        End If

        Dim ls_Dir As New IO.DirectoryInfo(ls_LogFolder & "\Log\Process\")
        Dim ls_GetFile As IO.FileInfo() = ls_Dir.GetFiles()
        Dim ls_File As IO.FileInfo
        Dim li_index As Long
        Dim ls_log As String
        Dim li_CountDate As Long
        Dim ls_Temp As List(Of String) = New List(Of String)

        li_index = 0
        For Each ls_File In ls_GetFile

            ls_Temp.Add(ls_File.ToString)

            ls_log = Right(ls_Temp.Item(li_index), 12)
            ls_log = Mid(ls_log, 1, 8)
            ls_log = Left(ls_log, 4) & "/" & Mid(ls_log, 5, 2) & "/" & Right(ls_log, 2)
            li_CountDate = DateDiff(DateInterval.Day, CDate(Format(CDate(ls_log), "yyyy/MM/dd")), CDate(Format(Now, "yyyy/MM/dd")))
            If li_CountDate > 90 Then
                File.Delete(ls_LogFolder & "\Log\Process\prc" & pScreenName & "_" & Format(CDate(ls_log), "yyyyMMdd") & ".log")
            End If

            li_index = li_index + 1
        Next
    End Sub


    '*************************************************************
    'NAME:          WriteToProcessLog
    'PURPOSE:       Open or create an error log and submit error message
    'PARAMETERS:    pStartTime      - Start of time process of the error file entry
    '               pEndTime        - End of time process of the error file entry
    '               pScreenName     - Name of Screen Process
    '               pCustomMsg      - Name of Process of the error file entry
    '               pErrSummary     - Err summary process
    '               pErrID          - Error ID
    '               pErrSeverity    - Enum ALERT, INFO and Err of the error file entry
    '               pWriteEventLog  - True false for Write to Event Log
    '               pStartEndStatus - Different Format for start and end
    '               pUseLogTime     - True or False use log time
    'RETURNS:       Nothing
    '*************************************************************
    ''' <summary>
    ''' this procedure for process error log
    ''' </summary>
    ''' <param name="pStartTime">Start Time</param>
    ''' <param name="pEndTime">End Time</param>
    ''' <param name="pScreenName">Screen Name</param>
    ''' <param name="pCustomMsg">Message</param>
    ''' <param name="pErrSummary">Err Summary</param>
    ''' <param name="pErrID">Err ID</param>
    ''' <param name="pErrSeverity">Err Severity</param>
    ''' <param name="pWriteToEventLog">Write to Event Log</param>
    ''' <param name="pStartEndStatus">Start or End Status</param>
    ''' <param name="pUseLogTime">Use Log Time</param>
    ''' <remarks>
    ''' 1. create \Log\ directory if not create before
    ''' 2. create \Log\Process\ directory if not create before
    ''' 3. checking the process log file with the same date if found it will be add the process information with the exist file
    ''' 4. create the process detail information
    ''' 5. keep the process log file from last 90 days until now otherwise will removed
    ''' </remarks>
    Public Sub WriteToProcessLog(ByVal pStartTime As Date, ByVal pEndTime As Date, ByVal pScreenName As String,
                                 Optional ByVal pCustomMsg As String = "",
                                 Optional ByVal pErrSummary As String = "",
                                 Optional ByVal pErrID As Integer = 9999,
                                 Optional ByVal pErrSeverity As ErrSeverity = ErrSeverity.INFO,
                                 Optional ByVal pWriteToEventLog As Boolean = False,
                                 Optional ByVal pStartEndStatus As String = "",
                                 Optional ByVal pUseLogTime As Boolean = False,
                                 Optional ByVal pLogID As String = "",
                                 Optional ByVal pParameters As String = "")

        Dim ls_Date As String
        Dim ls_ErrType As String
        Dim ls_Duration As String
        Dim ls_DurationError As String
        Dim ls_datefolder As String
        Dim ls_CompName As String = uf_CompName()
        Dim ls_LogFolder As String = "D:\RelayPLC"


        If Not System.IO.Directory.Exists("D:\") Then
            ls_LogFolder = Application.StartupPath
        End If


        li_StartTime = pStartTime
        li_EndTime = pEndTime
        li_Duration = li_EndTime - li_StartTime
        ls_Duration = uf_AddSpace(Format(li_Duration.TotalMilliseconds, "###.#0") & " (ms)", 15)
        ls_DurationError = uf_AddSpace("", 15)
        ls_datefolder = Format(Now, "yyyyMMdd")
        ls_Date = Format(Now, "yyyyMMdd")
        ls_ErrType = uf_ErrSeverity(pErrSeverity)
        pScreenName = pScreenName.Trim
        pScreenName = pScreenName.Replace(" ", "_")

        If Not System.IO.Directory.Exists(ls_LogFolder) Then
            System.IO.Directory.CreateDirectory(ls_LogFolder)
        End If

        If Not System.IO.Directory.Exists(ls_LogFolder & "\Log") Then
            System.IO.Directory.CreateDirectory(ls_LogFolder & "\Log")
        End If

        If Not System.IO.Directory.Exists(ls_LogFolder &
        "\Log\" & ls_datefolder) Then
            System.IO.Directory.CreateDirectory(ls_LogFolder &
            "\Log\" & ls_datefolder)
        End If


        If Not System.IO.Directory.Exists(ls_LogFolder &
        "\Log\" & ls_datefolder & "\Process") Then
            System.IO.Directory.CreateDirectory(ls_LogFolder &
            "\Log\" & ls_datefolder & "\Process")
        End If

        'check the file
        Dim fs As FileStream = New FileStream(ls_LogFolder &
        "\Log\" & ls_datefolder & "\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.OpenOrCreate, FileAccess.ReadWrite)


        Dim s As StreamWriter = New StreamWriter(fs)

        s.Close()
        fs.Close()

        'log it
        Dim fs1 As FileStream = New FileStream(ls_LogFolder &
        "\Log\" & ls_datefolder & "\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)


        If pStartEndStatus.ToUpper.Trim = "START" Then
            If pUseLogTime = True Then
                s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
            Else
                s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
            End If
            s1.Write("[" & UserLogin & "] ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & " ")
            s1.Write("[" & ls_ErrType & "] ")
            s1.Write("start.")
            s1.Write("" & vbCrLf)
            s1.Write("* * * * *")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        ElseIf pStartEndStatus.ToUpper.Trim = "END" Then
            s1.Write("* * * * *")
            s1.Write("" & vbCrLf)
            If pUseLogTime = True Then
                s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
            Else
                s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
            End If
            s1.Write("[" & UserLogin & "] ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & " ")
            s1.Write("[" & ls_ErrType & "] ")
            s1.Write("end.")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        ElseIf pStartEndStatus = "" Then
            If ls_ErrType = "I" Then
                If pUseLogTime = True Then
                    s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
                Else
                    s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
                End If
                s1.Write("[" & UserLogin & "] ")
                s1.Write("[" & ls_CompName & "] ")
                s1.Write("" & pScreenName & " ")
                s1.Write("[" & ls_ErrType & "] ")
                s1.Write("" & ls_Duration & " ")
                s1.Write("" & pCustomMsg & " ")
                s1.Write("" & pErrSummary & "")
                s1.Write("" & vbCrLf)
                s1.Close()
                fs1.Close()
            Else
                If pUseLogTime = True Then
                    s1.Write("" & Format(CDate(Now), "dd/MM/yyyy HH:mm:ss") & " ")
                Else
                    s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
                End If
                s1.Write("[" & UserLogin & "] ")
                s1.Write("[" & ls_CompName & "] ")
                s1.Write("" & pScreenName & " ")
                s1.Write("[" & ls_ErrType & "] ")
                s1.Write("" & ls_DurationError & " ")
                s1.Write("" & pCustomMsg & " ")
                s1.Write("" & pErrSummary & "")
                s1.Write("" & vbCrLf)
                s1.Close()
                fs1.Close()
            End If
        End If

        'write to event log
        If pWriteToEventLog = True Then
            If ls_ErrType = "E" Then
                WriteToEventLog(pCustomMsg, "RelayPLC(" & pScreenName & ")", EventLogEntryType.Error, "Application", pErrID)
            ElseIf ls_ErrType = "A" Then
                WriteToEventLog(pCustomMsg, "RelayPLC(" & pScreenName & ")", EventLogEntryType.Warning, "Application", pErrID)
            End If
        End If

        'Dim ls_Dir As New IO.DirectoryInfo(ls_LogFolder & "\Log\Process\")
        'Dim ls_GetFile As IO.FileInfo() = ls_Dir.GetFiles()
        'Dim ls_File As IO.FileInfo
        'Dim li_index As Long
        'Dim ls_log As String
        'Dim li_CountDate As Long
        'Dim ls_Temp As List(Of String) = New List(Of String)

        'li_index = 0
        'For Each ls_File In ls_GetFile

        '    ls_Temp.Add(ls_File.ToString)

        '    ls_log = Right(ls_Temp.Item(li_index), 12)
        '    ls_log = Mid(ls_log, 1, 8)
        '    ls_log = Mid(ls_log, 5, 2) & "/" & Right(ls_log, 2) & "/" & Left(ls_log, 4)
        '    li_CountDate = DateDiff(DateInterval.Day, CDate(Format(CDate(ls_log), "MM/dd/yyyy")), CDate(Format(Now, "MM/dd/yyyy")))
        '    If li_CountDate > 90 Then
        '        File.Delete(ls_LogFolder & "\Log\Process\prc" & pScreenName & "_" & Format(CDate(ls_log), "yyyyMMdd") & ".log")
        '    End If

        '    li_index = li_index + 1
        'Next
    End Sub

    '*************************************************************
    'NAME:          WriteToProcessLog
    'PURPOSE:       Open or create an error log and submit error message
    'PARAMETERS:    pStartTime - Start of time process of the error file entry
    '               pScreenName - Name of Screen Process
    '               pCustomMsg - Name of Process of the error file entry
    '               pErrSummary - Err summary process
    '               pErrSeverity - Enum ALERT, INFO and Err of the error file entry
    '               pErrMsg - msg of the error file entry
    'RETURNS:       Nothing
    '*************************************************************
    ''' <summary>
    ''' this procedure for process error log
    ''' </summary>
    ''' <param name="pStartTime">Start Time</param>
    ''' <param name="pScreenName">Screen Name</param>
    ''' <param name="pLogID">LogID</param>
    ''' <param name="pParameters">Parameters Message</param>
    ''' <remarks>
    ''' 1. create \Log\ directory if not create before
    ''' 2. create \Log\Process\ directory if not create before
    ''' 3. checking the process log file with the same date if found it will be add the process information with the exist file
    ''' 4. create the process detail information
    ''' 5. keep the process log file from last 90 days until now otherwise will removed
    ''' </remarks>
    Public Sub WriteToProcessLog(ByVal pLogID As String, ByVal pStartTime As Date, ByVal pScreenName As String,
                                 Optional ByVal pParameters As String = "",
                                 Optional ByVal pHarigami As Boolean = False)

        'check and make the directory if necessary; this is set to look in 
        'the application folder, you may wish to place the error log in 
        'another location depending upon the user's role and write access to 
        'different areas of the file system
        Dim ls_Date As String
        Dim ls_ErrType As String
        Dim ls_Duration As String
        Dim ls_CustomMsg As String
        Dim li_EventID As Integer
        Dim ls_LogFolder As String = "D:\RelayPLC"
        Dim Message As String
        Dim MsgFound As Boolean = False
        Dim ls_DateFolder As String = ""
        Dim ls_CompName As String = uf_CompName()
        '========================================================================
        'Message
        '========================================================================

        If Not System.IO.Directory.Exists("D:\") Then
            ls_LogFolder = Application.StartupPath
        End If

        Message = uf_ConvertMsg(pLogID)

        If Message = "" Then
            Message = pLogID
        Else
            MsgFound = True
        End If
        Dim i As Integer, Position As Long
        Dim Parameters() As String
        Parameters = Split(pParameters, "|")
        If UBound(Parameters) <> -1 Then
            Position = InStr(1, Message, "%%")
            Do While Position > 0
                Message = Left(Message, Position - 1) & Parameters(i) & Mid(Message, Position + 2, Len(Message) - Position)
                Position = InStr(1, Message, "%%")
                i = i + 1
            Loop
        Else
            Message = Replace(Message, "%", "")
        End If
        If MsgFound Then
            ls_CustomMsg = Message
        Else
            ls_CustomMsg = Message
        End If

        li_EventID = uf_ConvertEventID(pLogID)

        li_StartTime = pStartTime
        li_EndTime = Now
        li_Duration = li_EndTime - li_StartTime
        ls_Duration = uf_AddSpace(Format(li_Duration.TotalMilliseconds, "###.#0") & " (ms)", 15)

        ls_Date = Format(Now, "yyyyMMdd")
        ls_DateFolder = ls_Date
        ls_ErrType = uf_ErrSeverity(ErrSeverity.ERR)
        pScreenName = pScreenName.Trim
        pScreenName = pScreenName.Replace(" ", "_")


        If pHarigami Then
            If Not System.IO.Directory.Exists(ls_LogFolder) Then
                System.IO.Directory.CreateDirectory(ls_LogFolder)
            End If

            If Not System.IO.Directory.Exists(ls_LogFolder & "\Log") Then
                System.IO.Directory.CreateDirectory(ls_LogFolder & "\Log")
            End If

            If Not System.IO.Directory.Exists(ls_LogFolder &
                    "\Log\" & ls_DateFolder) Then
                System.IO.Directory.CreateDirectory(ls_LogFolder &
                "\Log\" & ls_DateFolder)
            End If


            If Not System.IO.Directory.Exists(ls_LogFolder &
                    "\Log\" & ls_DateFolder & "\Process") Then
                System.IO.Directory.CreateDirectory(ls_LogFolder &
                "\Log\" & ls_DateFolder & "\Process")
            End If

            'check the file
            Dim fs As FileStream = New FileStream(ls_LogFolder &
            "\Log\" & ls_DateFolder & "\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.OpenOrCreate, FileAccess.ReadWrite)


            Dim s As StreamWriter = New StreamWriter(fs)

            s.Close()
            fs.Close()

            'log it
            Dim fs1 As FileStream = New FileStream(ls_LogFolder &
            "\Log\" & ls_DateFolder & "\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.Append, FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)

            s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
            s1.Write("[" & UserLogin & "] ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & " ")
            s1.Write("[" & ls_ErrType & "] ")
            s1.Write("" & ls_CustomMsg.Trim & "")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        Else
            If Not System.IO.Directory.Exists(ls_LogFolder) Then
                System.IO.Directory.CreateDirectory(ls_LogFolder)
            End If

            If Not System.IO.Directory.Exists(ls_LogFolder &
            "\Log\Process") Then
                System.IO.Directory.CreateDirectory(ls_LogFolder &
                "\Log\Process")
            End If

            'check the file
            Dim fs As FileStream = New FileStream(ls_LogFolder &
            "\Log\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.OpenOrCreate, FileAccess.ReadWrite)
            Dim s As StreamWriter = New StreamWriter(fs)

            s.Close()
            fs.Close()

            'log it
            Dim fs1 As FileStream = New FileStream(ls_LogFolder &
            "\Log\Process\prc" & pScreenName & "_" & ls_Date & ".log", FileMode.Append, FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)

            s1.Write("" & Format(CDate(li_StartTime), "dd/MM/yyyy HH:mm:ss") & " ")
            s1.Write("[" & UserLogin & "] ")
            s1.Write("[" & ls_CompName & "] ")
            s1.Write("" & pScreenName & " ")
            s1.Write("[" & ls_ErrType & "] ")
            s1.Write("" & ls_CustomMsg.Trim & "")
            s1.Write("" & vbCrLf)
            s1.Close()
            fs1.Close()
        End If




        'write to event log
        WriteToEventLog(ls_CustomMsg, "RelayPLC(" & pScreenName & ")", EventLogEntryType.Error, "Application", li_EventID)


        If Not pHarigami Then
            Dim ls_Dir As New IO.DirectoryInfo(ls_LogFolder & "\Log\Process\")
            Dim ls_GetFile As IO.FileInfo() = ls_Dir.GetFiles()
            Dim ls_File As IO.FileInfo
            Dim li_index As Long
            Dim ls_log As String
            Dim li_CountDate As Long
            Dim ls_Temp As List(Of String) = New List(Of String)

            li_index = 0
            For Each ls_File In ls_GetFile

                ls_Temp.Add(ls_File.ToString)

                ls_log = Right(ls_Temp.Item(li_index), 12)
                ls_log = Mid(ls_log, 1, 8)
                ls_log = Left(ls_log, 4) & "/" & Mid(ls_log, 5, 2) & "/" & Right(ls_log, 2)
                li_CountDate = DateDiff(DateInterval.Day, CDate(Format(CDate(ls_log), "yyyy/MM/dd")), CDate(Format(Now, "yyyy/MM/dd")))
                If li_CountDate > 90 Then
                    File.Delete(ls_LogFolder & "\Log\Process\prc" & pScreenName & "_" & Format(CDate(ls_log), "yyyyMMdd") & ".log")
                End If

                li_index = li_index + 1
            Next
        End If

    End Sub


    ''' <summary>
    ''' this procedure for event log process
    ''' </summary>
    ''' <param name="Entry">Message</param>
    ''' <param name="AppName">Source Name</param>
    ''' <param name="EventType">Event</param>
    ''' <param name="LogName">Log Name</param>
    ''' <param name="pID">ID Number</param>
    ''' <returns>True if successful, false if not</returns>
    ''' <remarks></remarks>
    Private Function WriteToEventLog(ByVal Entry As String, Optional ByVal AppName As String = "RelayPLC",
        Optional ByVal EventType As EventLogEntryType = EventLogEntryType.Information,
        Optional ByVal LogName As String = "Application", Optional ByVal pID As Integer = 1000) As Boolean

        '*************************************************************
        'PURPOSE: Write Entry to Event Log using VB.NET
        'PARAMETERS: Entry - Value to Write
        '            AppName - Name of Client Application. Needed 
        '              because before writing to event log, you must 
        '              have a named EventLog source. 
        '            EventType - Entry Type, from EventLogEntryType 
        '              Structure e.g., EventLogEntryType.Warning, 
        '              EventLogEntryType.Error
        '            LogName: Name of Log (System, Application; 
        '              Security is read-only) If you 
        '              specify a non-existent log, the log will be
        '              created

        'RETURNS:   True if successful, false if not

        'EXAMPLES: 
        '1. Simple Example, Accepting All Defaults
        '    WriteToEventLog "Hello Event Log"

        '2.  Specify EventSource, EventType, and LogName
        '    WriteToEventLog("Danger, Danger, Danger", "MyVbApp", _
        '                      EventLogEntryType.Warning, "System")
        '
        'NOTE:     EventSources are tightly tied to their log. 
        '          So don't use the same source name for different 
        '          logs, and vice versa
        '******************************************************

        Dim objEventLog As New EventLog()

        Try
            'Register the App as an Event Source
            If Not EventLog.SourceExists(AppName) Then
                EventLog.CreateEventSource(AppName, LogName)
            End If

            objEventLog.Source = AppName

            'WriteEntry is overloaded; this is one
            'of 10 ways to call it
            objEventLog.WriteEntry(Entry, EventType, pID, CShort(0))
            Return True
        Catch Ex As Exception

            Return False

        End Try

    End Function

    Private Function uf_ErrSeverity(ByVal pErrSeverity As ErrSeverity) As String
        If pErrSeverity = ErrSeverity.ALERT Then
            Return "A"
        ElseIf pErrSeverity = ErrSeverity.ERR Then
            Return "E"
        Else
            Return "I"
        End If
    End Function

    Private Function uf_AddSpace(ByVal pDuration As String, ByVal pSpace As Integer) As String
        Return Space(pSpace - pDuration.Length) & pDuration
    End Function

    Private Function uf_CompName() As String
        Return My.Computer.Name.ToString
    End Function

#End Region

End Class
