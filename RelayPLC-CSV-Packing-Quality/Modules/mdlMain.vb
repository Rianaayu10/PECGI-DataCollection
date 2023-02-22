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

    Public gs_ServerPathShrinkCutting1 As String = ""
    Public gs_LocalPathShrinkCutting1 As String = ""
    Public gs_UserShrinkCutting1 As String = ""
    Public gs_PasswordShrinkCutting1 As String = ""
    Public gs_FileName_QualityShrinkCutting1 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalShrinkCutting1 As Integer = All_Interval

    Public gs_ServerPathWeightBatch1 As String = ""
    Public gs_LocalPathWeightBatch1 As String = ""
    Public gs_UserWeightBatch1 As String = ""
    Public gs_PasswordWeightBatch1 As String = ""
    Public gs_FileName_QualityWeightBatch1 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalWeightBatch1 As Integer = All_Interval

    Public gs_ServerPathHeater1 As String = ""
    Public gs_LocalPathHeater1 As String = ""
    Public gs_UserHeater1 As String = ""
    Public gs_PasswordHeater1 As String = ""
    Public gs_FileName_QualityHeater1 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalHeater1 As Integer = All_Interval

    Public gs_ServerPathTaping1 As String = ""
    Public gs_LocalPathTaping1 As String = ""
    Public gs_UserTaping1 As String = ""
    Public gs_PasswordTaping1 As String = ""
    Public gs_FileName_QualityTaping1 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalTaping1 As Integer = All_Interval

    Public gs_ServerPathWeightBox1 As String = ""
    Public gs_LocalPathWeightBox1 As String = ""
    Public gs_UserWeightBox1 As String = ""
    Public gs_PasswordWeightBox1 As String = ""
    Public gs_FileName_QualityWeightBox1 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalWeightBox1 As Integer = All_Interval

    Public gs_ServerPathShrinkCutting2 As String = ""
    Public gs_LocalPathShrinkCutting2 As String = ""
    Public gs_UserShrinkCutting2 As String = ""
    Public gs_PasswordShrinkCutting2 As String = ""
    Public gs_FileName_QualityShrinkCutting2 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalShrinkCutting2 As Integer = All_Interval

    Public gs_ServerPathWeightBatch2 As String = ""
    Public gs_LocalPathWeightBatch2 As String = ""
    Public gs_UserWeightBatch2 As String = ""
    Public gs_PasswordWeightBatch2 As String = ""
    Public gs_FileName_QualityWeightBatch2 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalWeightBatch2 As Integer = All_Interval

    Public gs_ServerPathHeater2 As String = ""
    Public gs_LocalPathHeater2 As String = ""
    Public gs_UserHeater2 As String = ""
    Public gs_PasswordHeater2 As String = ""
    Public gs_FileName_QualityHeater2 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalHeater2 As Integer = All_Interval

    Public gs_ServerPathTaping2 As String = ""
    Public gs_LocalPathTaping2 As String = ""
    Public gs_UserTaping2 As String = ""
    Public gs_PasswordTaping2 As String = ""
    Public gs_FileName_QualityTaping2 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalTaping2 As Integer = All_Interval

    Public gs_ServerPathWeightBox2 As String = ""
    Public gs_LocalPathWeightBox2 As String = ""
    Public gs_UserWeightBox2 As String = ""
    Public gs_PasswordWeightBox2 As String = ""
    Public gs_FileName_QualityWeightBox2 As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalWeightBox2 As Integer = All_Interval

    Public gs_ServerPathPillowProcess As String = ""
    Public gs_LocalPathPillowProcess As String = ""
    Public gs_UserPillowProcess As String = ""
    Public gs_PasswordPillowProcess As String = ""
    Public gs_FileName_QualityPillowProcess As String = "LOG_ProcessInfo.csv"
    Public gi_IntervalPillowProcess As Integer = All_Interval

    Public ConStr As String = setting.ConnectionString
    Public list() As String = {"1:16", "17:32", "33:48", "49:64", "65:80", "81:96", "97:112", "113:128", "129:144", "145:160"}

#End Region

End Module
