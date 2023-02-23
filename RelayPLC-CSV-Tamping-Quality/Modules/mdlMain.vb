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
    Public gs_FileName_Quality_HighSpeedMixerMch As String = "LOG31004_ProcessInfo.csv"
    Public gi_Interval_HighSpeedMixerMch As Integer = All_Interval

    Public ConStr As String = setting.ConnectionString
    Public list() As String = {"1:16", "17:32", "33:48", "49:64", "65:80", "81:96", "97:112", "113:128", "129:144", "145:160"}

#End Region

End Module
