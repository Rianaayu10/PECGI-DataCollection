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

    Public gs_ServerPathT05 As String = ""
    Public gs_LocalPathT05 As String = ""
    Public gs_UserT05 As String = ""
    Public gs_PasswordT05 As String = ""
    Public gs_FileName_HistoryT05 As String = "LOG11039_OperationHistory.csv"
    Public gi_IntervalT05 As Integer = All_Interval

    Public gs_ServerPathT06 As String = ""
    Public gs_LocalPathT06 As String = ""
    Public gs_UserT06 As String = ""
    Public gs_PasswordT06 As String = ""
    Public gs_FileName_HistoryT06 As String = "LOG11040_OperationHistory.csv"
    Public gi_IntervalT06 As Integer = All_Interval

    Public gs_ServerPathT07 As String = ""
    Public gs_LocalPathT07 As String = ""
    Public gs_UserT07 As String = ""
    Public gs_PasswordT07 As String = ""
    Public gs_FileName_HistoryT07 As String = "LOG11041_OperationHistory.csv"
    Public gi_IntervalT07 As Integer = All_Interval

    Public gs_ServerPathT08 As String = ""
    Public gs_LocalPathT08 As String = ""
    Public gs_UserT08 As String = ""
    Public gs_PasswordT08 As String = ""
    Public gs_FileName_HistoryT08 As String = "LOG11042_OperationHistory.csv"
    Public gi_IntervalT08 As Integer = All_Interval

    Public gs_ServerPathT09 As String = ""
    Public gs_LocalPathT09 As String = ""
    Public gs_UserT09 As String = ""
    Public gs_PasswordT09 As String = ""
    Public gs_FileName_HistoryT09 As String = "LOG11063_OperationHistory.csv"
    Public gi_IntervalT09 As Integer = All_Interval

    Public gs_ServerPathT10 As String = ""
    Public gs_LocalPathT10 As String = ""
    Public gs_UserT10 As String = ""
    Public gs_PasswordT10 As String = ""
    Public gs_FileName_HistoryT10 As String = "LOG11043_OperationHistory.csv"
    Public gi_IntervalT10 As Integer = All_Interval

    Public gs_ServerPathT12 As String = ""
    Public gs_LocalPathT12 As String = ""
    Public gs_UserT12 As String = ""
    Public gs_PasswordT12 As String = ""
    Public gs_FileName_HistoryT12 As String = "LOG11064_OperationHistory.csv"
    Public gi_IntervalT12 As Integer = All_Interval

    Public ConStr As String = setting.ConnectionString
    Public list() As String = {"1:16", "17:32", "33:48", "49:64", "65:80", "81:96", "97:112", "113:128", "129:144", "145:160"}

#End Region

End Module
