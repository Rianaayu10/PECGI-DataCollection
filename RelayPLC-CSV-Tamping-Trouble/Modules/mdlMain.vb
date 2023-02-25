Imports System.Data.SqlClient
Imports C1.Win.C1List.C1Combo
Imports System.Windows.Forms

Public Module mdlMain
#Region "Initial"
    Dim setting As clsConfig = New clsConfig

    Public sql As String
    Public gs_AppID As String = "P01"
    Public gs_Server As String = setting.Server
    Public gs_Database As String = setting.Database
    Public gs_User As String = setting.User
    Public gs_Password As String = setting.Password
    Public gi_RetryInterval As Integer = setting.RetryInterval
    Public gi_MaxRetry As Integer = setting.MaxRetry
    Public All_Interval As Integer = 1500

    Public Builder As SqlConnectionStringBuilder

    Public gs_ServerPath_HighSpeedMixerMch As String = ""
    Public gs_LocalPath_HighSpeedMixerMch As String = ""
    Public gs_User_HighSpeedMixerMch As String = ""
    Public gs_Password_HighSpeedMixerMch As String = ""
    Public gs_FileName_History_HighSpeedMixerMch As String = "LOG31004_OperationHistory.csv"
    Public gi_Interval_HighSpeedMixerMch As Integer = All_Interval

    Public gs_ServerPath_OscilatorMach As String = ""
    Public gs_LocalPath_OscilatorMach As String = ""
    Public gs_User_OscilatorMach As String = ""
    Public gs_Password_OscilatorMach As String = ""
    Public gs_FileName_History_OscilatorMach As String = "LOG31006_OperationHistory.csv"
    Public gi_Interval_OscilatorMach As Integer = All_Interval

    Public gs_ServerPath_Filling As String = ""
    Public gs_LocalPath_Filling As String = ""
    Public gs_User_Filling As String = ""
    Public gs_Password_Filling As String = ""
    Public gs_FileName_History_Filling As String = "LOG31010_OperationHistory.csv"
    Public gi_Interval_Filling As Integer = All_Interval

    Public gs_ServerPath_MesinPressing As String = ""
    Public gs_LocalPath_MesinPressing As String = ""
    Public gs_User_MesinPressing As String = ""
    Public gs_Password_MesinPressing As String = ""
    Public gs_FileName_History_MesinPressing As String = "LOG31016_OperationHistory.csv"
    Public gi_Interval_MesinPressing As Integer = All_Interval

    Public gs_ServerPath_SlitterMach1 As String = ""
    Public gs_LocalPath_SlitterMach1 As String = ""
    Public gs_User_SlitterMach1 As String = ""
    Public gs_Password_SlitterMach1 As String = ""
    Public gs_FileName_History_SlitterMach1 As String = "LOG31022_OperationHistory.csv"
    Public gi_Interval_SlitterMach1 As Integer = All_Interval

    Public gs_ServerPath_SlitterMach2 As String = ""
    Public gs_LocalPath_SlitterMach2 As String = ""
    Public gs_User_SlitterMach2 As String = ""
    Public gs_Password_SlitterMach2 As String = ""
    Public gs_FileName_History_SlitterMach2 As String = "LOG31023_OperationHistory.csv"
    Public gi_Interval_SlitterMach2 As Integer = All_Interval

    Public ConStr As String = setting.ConnectionString
    Public list() As String = {"1:16", "17:32", "33:48", "49:64", "65:80", "81:96", "97:112", "113:128", "129:144", "145:160"}

#End Region

End Module
