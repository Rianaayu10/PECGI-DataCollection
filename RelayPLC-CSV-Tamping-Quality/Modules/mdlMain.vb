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
    Public gs_FileName_QualityT05 As String = "LOG_OperationHistory.csv"
    Public gi_IntervalT05 As Integer = All_Interval

    Public gs_ServerPathT06 As String = ""
    Public gs_LocalPathT06 As String = ""
    Public gs_UserT06 As String = ""
    Public gs_PasswordT06 As String = ""
    Public gs_FileName_QualityT06 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalT06 As Integer = All_Interval

    Public gs_ServerPathT07 As String = ""
    Public gs_LocalPathT07 As String = ""
    Public gs_UserT07 As String = ""
    Public gs_PasswordT07 As String = ""
    Public gs_FileName_QualityT07 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalT07 As Integer = All_Interval

    Public gs_ServerPathT08 As String = ""
    Public gs_LocalPathT08 As String = ""
    Public gs_UserT08 As String = ""
    Public gs_PasswordT08 As String = ""
    Public gs_FileName_QualityT08 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalT08 As Integer = All_Interval

    Public gs_ServerPathT09 As String = ""
    Public gs_LocalPathT09 As String = ""
    Public gs_UserT09 As String = ""
    Public gs_PasswordT09 As String = ""
    Public gs_FileName_QualityT09 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalT09 As Integer = All_Interval

    Public gs_ServerPathT10 As String = ""
    Public gs_LocalPathT10 As String = ""
    Public gs_UserT10 As String = ""
    Public gs_PasswordT10 As String = ""
    Public gs_FileName_QualityT10 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalT10 As Integer = All_Interval

    Public gs_ServerPathT12 As String = ""
    Public gs_LocalPathT12 As String = ""
    Public gs_UserT12 As String = ""
    Public gs_PasswordT12 As String = ""
    Public gs_FileName_QualityT12 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalT12 As Integer = All_Interval

    Public gs_ServerPathOvenSeparator As String = ""
    Public gs_LocalPathOvenSeparator As String = ""
    Public gs_UserOvenSeparator As String = "OSP"
    Public gs_PasswordOvenSeparator As String = ""
    Public gs_FileName_QualityOvenSeparator As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalOvenSeparator As Integer = All_Interval

    Public gs_ServerPathOvenBobin As String = ""
    Public gs_LocalPathOvenBobin As String = ""
    Public gs_UserOvenBobin As String = "OBO"
    Public gs_PasswordOvenBobin As String = ""
    Public gs_FileName_QualityOvenBobin As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalOvenBobin As Integer = All_Interval

    Public gs_ServerPathOvenBubuk As String = ""
    Public gs_LocalPathOvenBubuk As String = ""
    Public gs_UserOvenBubuk As String = "OBU"
    Public gs_PasswordOvenBubuk As String = ""
    Public gs_FileName_QualityOvenBubuk As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalOvenBubuk As Integer = All_Interval

    Public gs_ServerPathOvenMangan As String = ""
    Public gs_LocalPathOvenMangan As String = ""
    Public gs_UserOvenMangan As String = "OBU"
    Public gs_PasswordOvenMangan As String = ""
    Public gs_FileName_QualityOvenMangan As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalOvenMangan As Integer = All_Interval

    Public gs_ServerPathOvenCaseCarbon As String = ""
    Public gs_LocalPathOvenCaseCarbon As String = ""
    Public gs_UserOvenCaseCarbon As String = "OBU"
    Public gs_PasswordOvenCaseCarbon As String = ""
    Public gs_FileName_QualityOvenCaseCarbon As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalOvenCaseCarbon As Integer = All_Interval

    Public gs_ServerPathOvenAnodeCap As String = ""
    Public gs_LocalPathOvenAnodeCap As String = ""
    Public gs_UserOvenAnodeCap As String = "OBU"
    Public gs_PasswordOvenAnodeCap As String = ""
    Public gs_FileName_QualityOvenAnodeCap As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalOvenAnodeCap As Integer = All_Interval

    Public gs_ServerPathHighSpeedMixerMch As String = ""
    Public gs_LocalPathHighSpeedMixerMch As String = ""
    Public gs_UserHighSpeedMixerMch As String = ""
    Public gs_PasswordHighSpeedMixerMch As String = ""
    Public gs_FileName_QualityHighSpeedMixerMch As String = "LOG31004_ProcessInfo.csv"
    Public gi_IntervalHighSpeedMixerMch As Integer = All_Interval

    Public ConStr As String = setting.ConnectionString
    Public list() As String = {"1:16", "17:32", "33:48", "49:64", "65:80", "81:96", "97:112", "113:128", "129:144", "145:160"}

#End Region

End Module
