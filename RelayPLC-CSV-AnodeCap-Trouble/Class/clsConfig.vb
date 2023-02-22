
Imports System.Data.SqlClient
Imports System.Xml
Imports System.IO

Public Class clsConfig

#Region "Initial"
    Private builder As SqlConnectionStringBuilder
    Private builderSAP As SqlConnectionStringBuilder

    'Used for Alert System Database
    Private ls_path As String
    'Edit by Pebri
    Private strConPath As String

    Private cDESEncryption As New clsDESEncryption("TOS")

    Private m_Server As String
    Public Property Server As String
        Get
            Return m_Server
        End Get
        Set(ByVal value As String)
            m_Server = value
        End Set
    End Property

    Private m_Database As String
    Public Property Database As String
        Get
            Return m_Database
        End Get
        Set(ByVal value As String)
            m_Database = value
        End Set
    End Property

    Private m_User As String
    Public Property User As String
        Get
            Return m_User
        End Get
        Set(ByVal value As String)
            m_User = value
        End Set
    End Property

    Private m_Password As String
    Public Property Password As String
        Get
            Return m_Password
        End Get
        Set(ByVal value As String)
            m_Password = value
        End Set
    End Property

    Private m_Shift As String
    Public Property Shift As String
        Get
            Return m_Shift
        End Get
        Set(ByVal value As String)
            m_Shift = value
        End Set
    End Property

    Private m_CommandTimeout As Integer
    Public Property CommandTimeout As Integer
        Get
            Return m_CommandTimeout
        End Get
        Set(ByVal value As Integer)
            m_CommandTimeout = value
        End Set
    End Property

    Private m_DatabaseTimeout As Integer
    Public Property DatabaseTimeout As Integer
        Get
            Return m_DatabaseTimeout
        End Get
        Set(ByVal value As Integer)
            m_DatabaseTimeout = value
        End Set
    End Property

    Private m_WinMode As String
    Public Property WinMode As String
        Get
            Return m_WinMode
        End Get
        Set(ByVal value As String)
            m_WinMode = value
        End Set
    End Property

    Private m_MaxRetry As Integer
    Public Property MaxRetry As Integer
        Get
            Return m_MaxRetry
        End Get
        Set(ByVal value As Integer)
            m_MaxRetry = value
        End Set
    End Property

    Private m_RetryInterval As Integer
    Public Property RetryInterval As Integer
        Get
            Return m_RetryInterval
        End Get
        Set(ByVal value As Integer)
            m_RetryInterval = value
        End Set
    End Property

    Private m_Interval As Integer
    Public Property Interval() As Integer
        Get
            Return m_Interval
        End Get
        Set(ByVal value As Integer)
            m_Interval = value
        End Set
    End Property

    Private m_ConnectionString As String
    Public ReadOnly Property ConnectionString As String
        Get
            Return m_ConnectionString
        End Get
    End Property

    Private m_ServerPathBA As String
    Public Property ServerPathBA As String
        Get
            Return m_ServerPathBA
        End Get
        Set(ByVal value As String)
            m_ServerPathBA = value
        End Set
    End Property

    Private m_LocalPathBA As String
    Public Property LocalPathBA As String
        Get
            Return m_LocalPathBA
        End Get
        Set(ByVal value As String)
            m_LocalPathBA = value
        End Set
    End Property

    Private m_UserBA As String
    Public Property UserBA As String
        Get
            Return m_UserBA
        End Get
        Set(ByVal value As String)
            m_UserBA = value
        End Set
    End Property

    Private m_PasswordBA As String
    Public Property PasswordBA As String
        Get
            Return m_PasswordBA
        End Get
        Set(ByVal value As String)
            m_PasswordBA = value
        End Set
    End Property

    Private m_ServerPathMIX As String
    Public Property ServerPathMIX As String
        Get
            Return m_ServerPathMIX
        End Get
        Set(ByVal value As String)
            m_ServerPathMIX = value
        End Set
    End Property

    Private m_LocalPathMIX As String
    Public Property LocalPathMIX As String
        Get
            Return m_LocalPathMIX
        End Get
        Set(ByVal value As String)
            m_LocalPathMIX = value
        End Set
    End Property

    Private m_UserMIX As String
    Public Property UserMIX As String
        Get
            Return m_UserMIX
        End Get
        Set(ByVal value As String)
            m_UserMIX = value
        End Set
    End Property

    Private m_PasswordMIX As String
    Public Property PasswordMIX As String
        Get
            Return m_PasswordMIX
        End Get
        Set(ByVal value As String)
            m_PasswordMIX = value
        End Set
    End Property

    Private m_ServerPathKNE As String
    Public Property ServerPathKNE As String
        Get
            Return m_ServerPathKNE
        End Get
        Set(ByVal value As String)
            m_ServerPathKNE = value
        End Set
    End Property

    Private m_LocalPathKNE As String
    Public Property LocalPathKNE As String
        Get
            Return m_LocalPathKNE
        End Get
        Set(ByVal value As String)
            m_LocalPathKNE = value
        End Set
    End Property

    Private m_UserKNE As String
    Public Property UserKNE As String
        Get
            Return m_UserKNE
        End Get
        Set(ByVal value As String)
            m_UserKNE = value
        End Set
    End Property

    Private m_PasswordKNE As String
    Public Property PasswordKNE As String
        Get
            Return m_PasswordKNE
        End Get
        Set(ByVal value As String)
            m_PasswordKNE = value
        End Set
    End Property

    Private m_ServerPathOXI As String
    Public Property ServerPathOXI As String
        Get
            Return m_ServerPathOXI
        End Get
        Set(ByVal value As String)
            m_ServerPathOXI = value
        End Set
    End Property

    Private m_LocalPathOXI As String
    Public Property LocalPathOXI As String
        Get
            Return m_LocalPathOXI
        End Get
        Set(ByVal value As String)
            m_LocalPathOXI = value
        End Set
    End Property

    Private m_UserOXI As String
    Public Property UserOXI As String
        Get
            Return m_UserOXI
        End Get
        Set(ByVal value As String)
            m_UserOXI = value
        End Set
    End Property

    Private m_PasswordOXI As String
    Public Property PasswordOXI As String
        Get
            Return m_PasswordOXI
        End Get
        Set(ByVal value As String)
            m_PasswordOXI = value
        End Set
    End Property

    Private m_ServerPathIR As String
    Public Property ServerPathIR As String
        Get
            Return m_ServerPathIR
        End Get
        Set(ByVal value As String)
            m_ServerPathIR = value
        End Set
    End Property

    Private m_LocalPathIR As String
    Public Property LocalPathIR As String
        Get
            Return m_LocalPathIR
        End Get
        Set(ByVal value As String)
            m_LocalPathIR = value
        End Set
    End Property

    Private m_UserIR As String
    Public Property UserIR As String
        Get
            Return m_UserIR
        End Get
        Set(ByVal value As String)
            m_UserIR = value
        End Set
    End Property

    Private m_PasswordIR As String
    Public Property PasswordIR As String
        Get
            Return m_PasswordIR
        End Get
        Set(ByVal value As String)
            m_PasswordIR = value
        End Set
    End Property

    Private m_ServerPathTAM As String
    Public Property ServerPathTAM As String
        Get
            Return m_ServerPathTAM
        End Get
        Set(ByVal value As String)
            m_ServerPathTAM = value
        End Set
    End Property

    Private m_LocalPathTAM As String
    Public Property LocalPathTAM As String
        Get
            Return m_LocalPathTAM
        End Get
        Set(ByVal value As String)
            m_LocalPathTAM = value
        End Set
    End Property

    Private m_UserTAM As String
    Public Property UserTAM As String
        Get
            Return m_UserTAM
        End Get
        Set(ByVal value As String)
            m_UserTAM = value
        End Set
    End Property

    Private m_PasswordTAM As String
    Public Property PasswordTAM As String
        Get
            Return m_PasswordTAM
        End Get
        Set(ByVal value As String)
            m_PasswordTAM = value
        End Set
    End Property

#End Region

#Region "Method"
    Public Function AddSlash(ByVal Path As String) As String
        Dim Result As String = Path
        If Path.EndsWith("\") = False Then
            Result = Result + "\"
        End If
        Return Result
    End Function
#End Region

#Region "Controller"
    ''' <summary>
    ''' Open config file and store the value in local variables.
    ''' </summary>
    ''' <param name="pConfigFile"></param>
    ''' <remarks></remarks>
    Public Sub New(Optional ByVal pConfigFile As String = "config.xml")
        'Dim Ret As String = Space(1500)

        ls_path = AddSlash(My.Application.Info.DirectoryPath) & pConfigFile


        'strConPath = AddSlash(My.Application.Info.DirectoryPath) & "MySql_ConnectionString.txt"
        'Using reader As StreamReader = New StreamReader(strConPath)
        '    ' Read one line from file
        '    m_MySqlConnection = reader.ReadToEnd
        'End Using

        If Not My.Computer.FileSystem.FileExists(ls_path) Then
            Throw New Exception("Config file is not found!")
        End If

        'Check XML file is empty or not
        If Trim(IO.File.ReadAllText(ls_path).Length) = 0 Then Exit Sub

        'Load XML file
        Dim document = XDocument.Load(ls_path)

        ''''''''''''''''''''''''''''''''''01. Load Setting EZDB to Screen''''''''''''''''''''''''''''''''''
        Dim ASDB = document.Descendants("EZDB").FirstOrDefault()
        If Not IsNothing(ASDB) Then
            If Not IsNothing(ASDB.Element("ServerName")) Then m_Server = cDESEncryption.DecryptData(ASDB.Element("ServerName").Value)
            If Not IsNothing(ASDB.Element("Database")) Then m_Database = cDESEncryption.DecryptData(ASDB.Element("Database").Value)
            If Not IsNothing(ASDB.Element("UserID")) Then m_User = cDESEncryption.DecryptData(ASDB.Element("UserID").Value)
            If Not IsNothing(ASDB.Element("Password")) Then m_Password = cDESEncryption.DecryptData(ASDB.Element("Password").Value)
            If Not IsNothing(ASDB.Element("WinMode")) Then m_WinMode = cDESEncryption.DecryptData(ASDB.Element("WinMode").Value)
            If Not IsNothing(ASDB.Element("CommandTimeOut")) Then m_CommandTimeout = IIf(IsNumeric(cDESEncryption.DecryptData(ASDB.Element("CommandTimeOut").Value)) = True, cDESEncryption.DecryptData(ASDB.Element("CommandTimeOut").Value), 0)
            If Not IsNothing(ASDB.Element("DatabaseTimeOut")) Then m_DatabaseTimeout = IIf(IsNumeric(cDESEncryption.DecryptData(ASDB.Element("DatabaseTimeOut").Value)) = True, cDESEncryption.DecryptData(ASDB.Element("DatabaseTimeOut").Value), 0)
            If m_Server = "" Or m_Database = "" Then
                Throw New Exception("Application setting is empty")
            End If

            builder = New SqlConnectionStringBuilder
            builder.DataSource = m_Server
            builder.InitialCatalog = m_Database
            builder.UserID = m_User
            builder.Password = m_Password
            builder.IntegratedSecurity = m_WinMode = "win"

            m_ConnectionString = builder.ConnectionString

        End If
        ''''''''''''''''''''''''''''''''''END''''''''''''''''''''''''''''''''''        

        Dim PATHBA = document.Descendants("PATHBA").FirstOrDefault()

        If Not IsNothing(PATHBA) Then
            If Not IsNothing(PATHBA.Element("ServerPath")) Then m_ServerPathBA = cDESEncryption.DecryptData(PATHBA.Element("ServerPath").Value)
            If Not IsNothing(PATHBA.Element("LocalPath")) Then m_LocalPathBA = cDESEncryption.DecryptData(PATHBA.Element("LocalPath").Value)
            If Not IsNothing(PATHBA.Element("User")) Then m_UserBA = cDESEncryption.DecryptData(PATHBA.Element("User").Value)
            If Not IsNothing(PATHBA.Element("Password")) Then m_PasswordBA = cDESEncryption.DecryptData(PATHBA.Element("Password").Value)
        End If
    End Sub
#End Region

End Class
