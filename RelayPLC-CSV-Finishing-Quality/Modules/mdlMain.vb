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

    'CAP CHECKER Mac
    Public gs_ServerPath_CAPCHECKERMac As String = ""
    Public gs_LocalPath_CAPCHECKERMac As String = ""
    Public gs_User_CAPCHECKERMac As String = ""
    Public gs_Password_CAPCHECKERMac As String = ""
    Public gs_FileName_CAPCHECKERMac As String = "LOG34010_ProcessInfo.csv"
    Public gi_Interval_CAPCHECKERMac As Integer = All_Interval

    'CELL SUPPLY BATTERY Mach
    Public gs_ServerPath_CELLSUPPLYBATTERYMach As String = ""
    Public gs_LocalPath_CELLSUPPLYBATTERYMach As String = ""
    Public gs_User_CELLSUPPLYBATTERYMach As String = ""
    Public gs_Password_CELLSUPPLYBATTERYMach As String = ""
    Public gs_FileName_CELLSUPPLYBATTERYMach As String = "LOG34011_ProcessInfo.csv"
    Public gi_Interval_CELLSUPPLYBATTERYMach As Integer = All_Interval

    'RING INSERT Mach
    Public gs_ServerPath_RINGINSERTMach As String = ""
    Public gs_LocalPath_RINGINSERTMach As String = ""
    Public gs_User_RINGINSERTMach As String = ""
    Public gs_Password_RINGINSERTMach As String = ""
    Public gs_FileName_RINGINSERTMach As String = "LOG34012_ProcessInfo.csv"
    Public gi_Interval_RINGINSERTMach As Integer = All_Interval

    'TUBE INSERT Mach
    Public gs_ServerPath_TUBEINSERTMach As String = ""
    Public gs_LocalPath_TUBEINSERTMach As String = ""
    Public gs_User_TUBEINSERTMach As String = ""
    Public gs_Password_TUBEINSERTMach As String = ""
    Public gs_FileName_TUBEINSERTMach As String = "LOG31013_ProcessInfo.csv"
    Public gi_Interval_TUBEINSERTMach As Integer = All_Interval

    'TUBE SHRINK Mach
    Public gs_ServerPath_TUBESHRINKMach As String = ""
    Public gs_LocalPath_TUBESHRINKMach As String = ""
    Public gs_User_TUBESHRINKMach As String = ""
    Public gs_Password_TUBESHRINKMach As String = ""
    Public gs_FileName_TUBESHRINKMach As String = "LOG34014_ProcessInfo.csv"
    Public gi_Interval_TUBESHRINKMach As Integer = All_Interval

    'SIDE CHECKER Mach
    Public gs_ServerPath_SIDECHECKERMach As String = ""
    Public gs_LocalPath_SIDECHECKERMach As String = ""
    Public gs_User_SIDECHECKERMach As String = ""
    Public gs_Password_SIDECHECKERMach As String = ""
    Public gs_FileName_SIDECHECKERMach As String = "LOG34015_ProcessInfo.csv"
    Public gi_Interval_SIDECHECKERMach As Integer = All_Interval

    'CELL BOXING Mach
    Public gs_ServerPath_CELLBOXINGMach As String = ""
    Public gs_LocalPath_CELLBOXINGMach As String = ""
    Public gs_User_CELLBOXINGMach As String = ""
    Public gs_Password_CELLBOXINGMach As String = ""
    Public gs_FileName_CELLBOXINGMach As String = "LOG34016_ProcessInfo.csv"
    Public gi_Interval_CELLBOXINGMach As Integer = All_Interval

    'SUPPLY BOX MACH Mach
    Public gs_ServerPath_SUPPLYBOXMACHMach As String = ""
    Public gs_LocalPath_SUPPLYBOXMACHMach As String = ""
    Public gs_User_SUPPLYBOXMACHMach As String = ""
    Public gs_Password_SUPPLYBOXMACHMach As String = ""
    Public gs_FileName_SUPPLYBOXMACHMach As String = "LOG34017_ProcessInfo.csv"
    Public gi_Interval_SUPPLYBOXMACHMach As Integer = All_Interval

    'CELL SUPPLY Mach
    Public gs_ServerPath_CELLSUPPLYMach As String = ""
    Public gs_LocalPath_CELLSUPPLYMach As String = ""
    Public gs_User_CELLSUPPLYMach As String = ""
    Public gs_Password_CELLSUPPLYMach As String = ""
    Public gs_FileName_CELLSUPPLYMach As String = "LOG34018_ProcessInfo.csv"
    Public gi_Interval_CELLSUPPLYMach As Integer = All_Interval

    'UPPER RING INSERT Mach
    Public gs_ServerPath_UPPERRINGINSERTMach As String = ""
    Public gs_LocalPath_UPPERRINGINSERTMach As String = ""
    Public gs_User_UPPERRINGINSERTMach As String = ""
    Public gs_Password_UPPERRINGINSERTMach As String = ""
    Public gs_FileName_UPPERRINGINSERTMach As String = "LOG34019_ProcessInfo.csv"
    Public gi_Interval_UPPERRINGINSERTMach As Integer = All_Interval

    'LABELLER Mach
    Public gs_ServerPath_LABELLERMach As String = ""
    Public gs_LocalPath_LABELLERMach As String = ""
    Public gs_User_LABELLERMach As String = ""
    Public gs_Password_LABELLERMach As String = ""
    Public gs_FileName_LABELLERMach As String = "LOG34020_ProcessInfo.csv"
    Public gi_Interval_LABELLERMach As Integer = All_Interval

    'OCV INSPECTION Mach
    Public gs_ServerPath_OCVINSPECTIONMach As String = ""
    Public gs_LocalPath_OCVINSPECTIONMach As String = ""
    Public gs_User_OCVINSPECTIONMach As String = ""
    Public gs_Password_OCVINSPECTIONMach As String = ""
    Public gs_FileName_OCVINSPECTIONMach As String = "LOG34021_ProcessInfo.csv"
    Public gi_Interval_OCVINSPECTIONMach As Integer = All_Interval

    'SIDE INSPECTION Mach
    Public gs_ServerPath_SIDEINSPECTIONMach As String = ""
    Public gs_LocalPath_SIDEINSPECTIONMach As String = ""
    Public gs_User_SIDEINSPECTIONMach As String = ""
    Public gs_Password_SIDEINSPECTIONMach As String = ""
    Public gs_FileName_SIDEINSPECTIONMach As String = "LOG34022_ProcessInfo.csv"
    Public gi_Interval_SIDEINSPECTIONMach As Integer = All_Interval

    'CELL BOXING Mach 2
    Public gs_ServerPath_CELLBOXINGMach2 As String = ""
    Public gs_LocalPath_CELLBOXINGMach2 As String = ""
    Public gs_User_CELLBOXINGMach2 As String = ""
    Public gs_Password_CELLBOXINGMach2 As String = ""
    Public gs_FileName_CELLBOXINGMach2 As String = "LOG34023_ProcessInfo.csv"
    Public gi_Interval_CELLBOXINGMach2 As Integer = All_Interval

    Public ConStr As String = setting.ConnectionString
    Public list() As String = {"1:16", "17:32", "33:48", "49:64", "65:80", "81:96", "97:112", "113:128", "129:144", "145:160"}

#End Region

End Module
