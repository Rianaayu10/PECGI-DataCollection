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

    Public gs_ServerPath_MixingMachVacuumHT42H As String = ""
    Public gs_LocalPath_MixingMachVacuumHT42H As String = ""
    Public gs_User_MixingMachVacuumHT42H As String = ""
    Public gs_Password_MixingMachVacuumHT42H As String = ""
    Public gs_FileName_MixingMachVacuumHT42H As String = "LOG31001_ProcessInfo.csv"
    Public gi_Interval_MixingMachVacuumHT42H As Integer = All_Interval

    Public gs_ServerPath_MixingMachVacuumHT35H As String = ""
    Public gs_LocalPath_MixingMachVacuumHT35H As String = ""
    Public gs_User_MixingMachVacuumHT35H As String = ""
    Public gs_Password_MixingMachVacuumHT35H As String = ""
    Public gs_FileName_MixingMachVacuumHT35H As String = "LOG31002_ProcessInfo.csv"
    Public gi_Interval_MixingMachVacuumHT35H As Integer = All_Interval

    Public gs_ServerPath_MixingMachVacuumEC As String = ""
    Public gs_LocalPath_MixingMachVacuumEC As String = ""
    Public gs_User_MixingMachVacuumEC As String = ""
    Public gs_Password_MixingMachVacuumEC As String = ""
    Public gs_FileName_MixingMachVacuumEC As String = "LOG31003_ProcessInfo.csv"
    Public gi_Interval_MixingMachVacuumEC As Integer = All_Interval

    Public gs_ServerPath_HighSpeedMixerMch As String = ""
    Public gs_LocalPath_HighSpeedMixerMch As String = ""
    Public gs_User_HighSpeedMixerMch As String = ""
    Public gs_Password_HighSpeedMixerMch As String = ""
    Public gs_FileName_HighSpeedMixerMch As String = "LOG31004_ProcessInfo.csv"
    Public gi_Interval_HighSpeedMixerMch As Integer = All_Interval

    Public gs_ServerPath_Conveyortransferpowder As String = ""
    Public gs_LocalPath_Conveyortransferpowder As String = ""
    Public gs_User_Conveyortransferpowder As String = ""
    Public gs_Password_Conveyortransferpowder As String = ""
    Public gs_FileName_Conveyortransferpowder As String = "LOG31005_ProcessInfo.csv"
    Public gi_Interval_Conveyortransferpowder As Integer = All_Interval

    Public gs_ServerPath_OscilatorMach As String = ""
    Public gs_LocalPath_OscilatorMach As String = ""
    Public gs_User_OscilatorMach As String = ""
    Public gs_Password_OscilatorMach As String = ""
    Public gs_FileName_OscilatorMach As String = "LOG31006_ProcessInfo.csv"
    Public gi_Interval_OscilatorMach As Integer = All_Interval

    Public gs_ServerPath_Conveyortransferbox As String = ""
    Public gs_LocalPath_Conveyortransferbox As String = ""
    Public gs_User_Conveyortransferbox As String = ""
    Public gs_Password_Conveyortransferbox As String = ""
    Public gs_FileName_Conveyortransferbox As String = "LOG31007_ProcessInfo.csv"
    Public gi_Interval_Conveyortransferbox As Integer = All_Interval

    Public gs_ServerPath_Boxlift As String = ""
    Public gs_LocalPath_Boxlift As String = ""
    Public gs_User_Boxlift As String = ""
    Public gs_Password_Boxlift As String = ""
    Public gs_FileName_Boxlift As String = "LOG31008_ProcessInfo.csv"
    Public gi_Interval_Boxlift As Integer = All_Interval

    Public gs_ServerPath_Expandmetaltension1 As String = ""
    Public gs_LocalPath_Expandmetaltension1 As String = ""
    Public gs_User_Expandmetaltension1 As String = ""
    Public gs_Password_Expandmetaltension1 As String = ""
    Public gs_FileName_Expandmetaltension1 As String = "LOG31009_ProcessInfo.csv"
    Public gi_Interval_Expandmetaltension1 As Integer = All_Interval

    Public gs_ServerPath_Expandmetaltension2 As String = ""
    Public gs_LocalPath_Expandmetaltension2 As String = ""
    Public gs_User_Expandmetaltension2 As String = ""
    Public gs_Password_Expandmetaltension2 As String = ""
    Public gs_FileName_Expandmetaltension2 As String = "LOG31010_ProcessInfo.csv"
    Public gi_Interval_Expandmetaltension2 As Integer = All_Interval

    Public gs_ServerPath_OvenRollerMachManual1 As String = ""
    Public gs_LocalPath_OvenRollerMachManual1 As String = ""
    Public gs_User_OvenRollerMachManual1 As String = ""
    Public gs_Password_OvenRollerMachManual1 As String = ""
    Public gs_FileName_OvenRollerMachManual1 As String = "LOG31011_ProcessInfo.csv"
    Public gi_Interval_OvenRollerMachManual1 As Integer = All_Interval

    Public gs_ServerPath_OvenHeaterMachManual1 As String = ""
    Public gs_LocalPath_OvenHeaterMachManual1 As String = ""
    Public gs_User_OvenHeaterMachManual1 As String = ""
    Public gs_Password_OvenHeaterMachManual1 As String = ""
    Public gs_FileName_OvenHeaterMachManual1 As String = "LOG31012_ProcessInfo.csv"
    Public gi_Interval_OvenHeaterMachManual1 As Integer = All_Interval

    Public gs_ServerPath_OvenRollerMachManual2 As String = ""
    Public gs_LocalPath_OvenRollerMachManual2 As String = ""
    Public gs_User_OvenRollerMachManual2 As String = ""
    Public gs_Password_OvenRollerMachManual2 As String = ""
    Public gs_FileName_OvenRollerMachManual2 As String = "LOG31013_ProcessInfo.csv"
    Public gi_Interval_OvenRollerMachManual2 As Integer = All_Interval

    Public gs_ServerPath_OvenHeaterMachManual2 As String = ""
    Public gs_LocalPath_OvenHeaterMachManual2 As String = ""
    Public gs_User_OvenHeaterMachManual2 As String = ""
    Public gs_Password_OvenHeaterMachManual2 As String = ""
    Public gs_FileName_OvenHeaterMachManual2 As String = "LOG31014_ProcessInfo.csv"
    Public gi_Interval_OvenHeaterMachManual2 As Integer = All_Interval

    Public gs_ServerPath_HoopTension As String = ""
    Public gs_LocalPath_HoopTension As String = ""
    Public gs_User_HoopTension As String = ""
    Public gs_Password_HoopTension As String = ""
    Public gs_FileName_HoopTension As String = "LOG31015_ProcessInfo.csv"
    Public gi_Interval_HoopTension As Integer = All_Interval

    Public gs_ServerPath_MesinPressing As String = ""
    Public gs_LocalPath_MesinPressing As String = ""
    Public gs_User_MesinPressing As String = ""
    Public gs_Password_MesinPressing As String = ""
    Public gs_FileName_MesinPressing As String = "LOG31016_ProcessInfo.csv"
    Public gi_Interval_MesinPressing As Integer = All_Interval

    Public gs_ServerPath_RollerMachManual1 As String = ""
    Public gs_LocalPath_RollerMachManual1 As String = ""
    Public gs_User_RollerMachManual1 As String = ""
    Public gs_Password_RollerMachManual1 As String = ""
    Public gs_FileName_RollerMachManual1 As String = "LOG31017_ProcessInfo.csv"
    Public gi_Interval_RollerMachManual1 As Integer = All_Interval

    Public gs_ServerPath_RollerMachManual2 As String = ""
    Public gs_LocalPath_RollerMachManual2 As String = ""
    Public gs_User_RollerMachManual2 As String = ""
    Public gs_Password_RollerMachManual2 As String = ""
    Public gs_FileName_RollerMachManual2 As String = "LOG31018_ProcessInfo.csv"
    Public gi_Interval_RollerMachManual2 As Integer = All_Interval

    Public gs_ServerPath_RollerMachManual3 As String = ""
    Public gs_LocalPath_RollerMachManual3 As String = ""
    Public gs_User_RollerMachManual3 As String = ""
    Public gs_Password_RollerMachManual3 As String = ""
    Public gs_FileName_RollerMachManual3 As String = "LOG31019_ProcessInfo.csv"
    Public gi_Interval_RollerMachManual3 As Integer = All_Interval

    Public gs_ServerPath_RollerMachManual4 As String = ""
    Public gs_LocalPath_RollerMachManual4 As String = ""
    Public gs_User_RollerMachManual4 As String = ""
    Public gs_Password_RollerMachManual4 As String = ""
    Public gs_FileName_RollerMachManual4 As String = "LOG31020_ProcessInfo.csv"
    Public gi_Interval_RollerMachManual4 As Integer = All_Interval

    Public gs_ServerPath_RollerMachManual5 As String = ""
    Public gs_LocalPath_RollerMachManual5 As String = ""
    Public gs_User_RollerMachManual5 As String = ""
    Public gs_Password_RollerMachManual5 As String = ""
    Public gs_FileName_RollerMachManual5 As String = "LOG31021_ProcessInfo.csv"
    Public gi_Interval_RollerMachManual5 As Integer = All_Interval

    Public gs_ServerPath_SlitterMach1 As String = ""
    Public gs_LocalPath_SlitterMach1 As String = ""
    Public gs_User_SlitterMach1 As String = ""
    Public gs_Password_SlitterMach1 As String = ""
    Public gs_FileName_SlitterMach1 As String = "LOG31022_ProcessInfo.csv"
    Public gi_Interval_SlitterMach1 As Integer = All_Interval

    Public gs_ServerPath_SlitterMach2 As String = ""
    Public gs_LocalPath_SlitterMach2 As String = ""
    Public gs_User_SlitterMach2 As String = ""
    Public gs_Password_SlitterMach2 As String = ""
    Public gs_FileName_SlitterMach2 As String = "LOG31023_ProcessInfo.csv"
    Public gi_Interval_SlitterMach2 As Integer = All_Interval

    Public ConStr As String = setting.ConnectionString
    Public list() As String = {"1:16", "17:32", "33:48", "49:64", "65:80", "81:96", "97:112", "113:128", "129:144", "145:160"}

#End Region

End Module
