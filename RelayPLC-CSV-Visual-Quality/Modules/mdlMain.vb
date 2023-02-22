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

    Public gs_ServerPathWCH01 As String = ""
    Public gs_LocalPathWCH01 As String = ""
    Public gs_UserWCH01 As String = ""
    Public gs_PasswordWCH01 As String = ""
    Public gs_FileName_QualityWCH01 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalWCH01 As Integer = All_Interval

    Public gs_ServerPathAV01 As String = ""
    Public gs_LocalPathAV01 As String = ""
    Public gs_UserAV01 As String = ""
    Public gs_PasswordAV01 As String = ""
    Public gs_FileName_QualityAV01 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalAV01 As Integer = All_Interval

    Public gs_ServerPathAV02 As String = ""
    Public gs_LocalPathAV02 As String = ""
    Public gs_UserAV02 As String = ""
    Public gs_PasswordAV02 As String = ""
    Public gs_FileName_QualityAV02 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalAV02 As Integer = All_Interval

    Public gs_ServerPathM01 As String = ""
    Public gs_LocalPathM01 As String = ""
    Public gs_UserM01 As String = ""
    Public gs_PasswordM01 As String = ""
    Public gs_FileName_QualityM01 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalM01 As Integer = All_Interval

    Public gs_ServerPathIJ01 As String = ""
    Public gs_LocalPathIJ01 As String = ""
    Public gs_UserIJ01 As String = ""
    Public gs_PasswordIJ01 As String = ""
    Public gs_FileName_QualityIJ01 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalIJ01 As Integer = All_Interval

    Public gs_ServerPathIJ02 As String = ""
    Public gs_LocalPathIJ02 As String = ""
    Public gs_UserIJ02 As String = ""
    Public gs_PasswordIJ02 As String = ""
    Public gs_FileName_QualityIJ02 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalIJ02 As Integer = All_Interval
    Public ConStr As String = setting.ConnectionString

    Public list() As String = {"1:16", "17:32", "33:48", "49:64", "65:80", "81:96", "97:112", "113:128", "129:144", "145:160"}

#End Region

End Module
