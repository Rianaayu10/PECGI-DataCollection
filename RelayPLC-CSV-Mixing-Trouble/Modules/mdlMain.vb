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
    Public list() As String = {"1:16", "17:32", "33:48", "49:64", "65:80", "81:96", "97:112", "113:128", "129:144", "145:160"}

    Public gs_ServerPathMix01 As String = ""
    Public gs_LocalPathMix01 As String = ""
    Public gs_UserMix01 As String = ""
    Public gs_PasswordMix01 As String = ""
    Public gs_FileName_TroubleMix01 As String = "LOG11021_OperationHistory.csv"
    Public gi_IntervalMix01 As Integer = All_Interval

    Public gs_ServerPathMix02 As String = ""
    Public gs_LocalPathMix02 As String = ""
    Public gs_UserMix02 As String = ""
    Public gs_PasswordMix02 As String = ""
    Public gs_FileName_TroubleMix02 As String = "LOG11060_OperationHistory.csv"
    Public gi_IntervalMix02 As Integer = All_Interval

    Public gs_ServerPathMix03 As String = ""
    Public gs_LocalPathMix03 As String = ""
    Public gs_UserMix03 As String = ""
    Public gs_PasswordMix03 As String = ""
    Public gs_FileName_TroubleMix03 As String = "LOG11022_OperationHistory.csv"
    Public gi_IntervalMix03 As Integer = All_Interval

    Public gs_ServerPathMix04 As String = ""
    Public gs_LocalPathMix04 As String = ""
    Public gs_UserMix04 As String = ""
    Public gs_PasswordMix04 As String = ""
    Public gs_FileName_TroubleMix04 As String = "LOG11023_OperationHistory.csv"
    Public gi_IntervalMix04 As Integer = All_Interval

    Public gs_ServerPathMix05 As String = ""
    Public gs_LocalPathMix05 As String = ""
    Public gs_UserMix05 As String = ""
    Public gs_PasswordMix05 As String = ""
    Public gs_FileName_TroubleMix05 As String = "LOG11024_OperationHistory.csv"
    Public gi_IntervalMix05 As Integer = All_Interval

    Public ConStr As String = setting.ConnectionString
#End Region

End Module
