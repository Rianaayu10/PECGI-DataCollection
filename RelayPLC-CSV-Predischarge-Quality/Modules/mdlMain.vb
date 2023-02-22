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

    Public gs_ServerPathPD16 As String = ""
    Public gs_LocalPathPD16 As String = ""
    Public gs_UserPD16 As String = ""
    Public gs_PasswordPD16 As String = ""
    Public gs_FileName_QualityPD16 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalPD16 As Integer = All_Interval

    Public gs_ServerPathPD17 As String = ""
    Public gs_LocalPathPD17 As String = ""
    Public gs_UserPD17 As String = ""
    Public gs_PasswordPD17 As String = ""
    Public gs_FileName_QualityPD17 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalPD17 As Integer = All_Interval

    Public gs_ServerPathPD18 As String = ""
    Public gs_LocalPathPD18 As String = ""
    Public gs_UserPD18 As String = ""
    Public gs_PasswordPD18 As String = ""
    Public gs_FileName_QualityPD18 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalPD18 As Integer = All_Interval

    Public ConStr As String = setting.ConnectionString
#End Region

End Module
