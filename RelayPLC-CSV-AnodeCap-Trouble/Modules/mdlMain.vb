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
    Public gs_FileName_TroubleMix01 As String = "LOG12021_OperationHistory.csv"
    Public gi_IntervalMix01 As Integer = All_Interval

    Public gs_ServerPathMix02 As String = ""
    Public gs_LocalPathMix02 As String = ""
    Public gs_UserMix02 As String = ""
    Public gs_PasswordMix02 As String = ""
    Public gs_FileName_TroubleMix02 As String = "LOG12022_OperationHistory.csv"
    Public gi_IntervalMix02 As Integer = All_Interval

    Public gs_ServerPathMix03 As String = ""
    Public gs_LocalPathMix03 As String = ""
    Public gs_UserMix03 As String = ""
    Public gs_PasswordMix03 As String = ""
    Public gs_FileName_TroubleMix03 As String = "LOG12023_OperationHistory.csv"
    Public gi_IntervalMix03 As Integer = All_Interval

    Public gs_ServerPathMix04 As String = ""
    Public gs_LocalPathMix04 As String = ""
    Public gs_UserMix04 As String = ""
    Public gs_PasswordMix04 As String = ""
    Public gs_FileName_TroubleMix04 As String = "LOG12024_OperationHistory.csv"
    Public gi_IntervalMix04 As Integer = All_Interval

    Public gs_ServerPathMix05 As String = ""
    Public gs_LocalPathMix05 As String = ""
    Public gs_UserMix05 As String = ""
    Public gs_PasswordMix05 As String = ""
    Public gs_FileName_TroubleMix05 As String = "LOG12025_OperationHistory.csv"
    Public gi_IntervalMix05 As Integer = All_Interval

    Public gs_ServerPathMix06 As String = ""
    Public gs_LocalPathMix06 As String = ""
    Public gs_UserMix06 As String = ""
    Public gs_PasswordMix06 As String = ""
    Public gs_FileName_TroubleMix06 As String = "LOG12026_OperationHistory.csv"
    Public gi_IntervalMix06 As Integer = All_Interval

    Public gs_ServerPathMix07 As String = ""
    Public gs_LocalPathMix07 As String = ""
    Public gs_UserMix07 As String = ""
    Public gs_PasswordMix07 As String = ""
    Public gs_FileName_TroubleMix07 As String = "LOG12027_OperationHistory.csv"
    Public gi_IntervalMix07 As Integer = All_Interval

    Public gs_ServerPathMix08 As String = ""
    Public gs_LocalPathMix08 As String = ""
    Public gs_UserMix08 As String = ""
    Public gs_PasswordMix08 As String = ""
    Public gs_FileName_TroubleMix08 As String = "LOG12028_OperationHistory.csv"
    Public gi_IntervalMix08 As Integer = All_Interval

    Public gs_ServerPathMix09 As String = ""
    Public gs_LocalPathMix09 As String = ""
    Public gs_UserMix09 As String = ""
    Public gs_PasswordMix09 As String = ""
    Public gs_FileName_TroubleMix09 As String = "LOG12029_OperationHistory.csv"
    Public gi_IntervalMix09 As Integer = All_Interval

    Public gs_ServerPathMix10 As String = ""
    Public gs_LocalPathMix10 As String = ""
    Public gs_UserMix10 As String = ""
    Public gs_PasswordMix10 As String = ""
    Public gs_FileName_TroubleMix10 As String = "LOG12030_OperationHistory.csv"
    Public gi_IntervalMix10 As Integer = All_Interval

    Public gs_ServerPathMix11 As String = ""
    Public gs_LocalPathMix11 As String = ""
    Public gs_UserMix11 As String = ""
    Public gs_PasswordMix11 As String = ""
    Public gs_FileName_TroubleMix11 As String = "LOG12031_OperationHistory.csv"
    Public gi_IntervalMix11 As Integer = All_Interval

    Public gs_ServerPathMix12 As String = ""
    Public gs_LocalPathMix12 As String = ""
    Public gs_UserMix12 As String = ""
    Public gs_PasswordMix12 As String = ""
    Public gs_FileName_TroubleMix12 As String = "LOG12032_OperationHistory.csv"
    Public gi_IntervalMix12 As Integer = All_Interval

    Public gs_ServerPathMix13 As String = ""
    Public gs_LocalPathMix13 As String = ""
    Public gs_UserMix13 As String = ""
    Public gs_PasswordMix13 As String = ""
    Public gs_FileName_TroubleMix13 As String = "LOG12033_OperationHistory.csv"
    Public gi_IntervalMix13 As Integer = All_Interval

    Public gs_ServerPathMix14 As String = ""
    Public gs_LocalPathMix14 As String = ""
    Public gs_UserMix14 As String = ""
    Public gs_PasswordMix14 As String = ""
    Public gs_FileName_TroubleMix14 As String = "LOG12034_OperationHistory.csv"
    Public gi_IntervalMix14 As Integer = All_Interval

    Public gs_ServerPathMix15 As String = ""
    Public gs_LocalPathMix15 As String = ""
    Public gs_UserMix15 As String = ""
    Public gs_PasswordMix15 As String = ""
    Public gs_FileName_TroubleMix15 As String = "LOG12035_OperationHistory.csv"
    Public gi_IntervalMix15 As Integer = All_Interval

    Public ConStr As String = setting.ConnectionString
#End Region

End Module
