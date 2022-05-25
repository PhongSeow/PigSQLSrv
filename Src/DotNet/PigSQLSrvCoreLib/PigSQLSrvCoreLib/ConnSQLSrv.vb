'**********************************
'* Name: ConnSQLSrv
'* Author: Seow Phong
'* License: Copyright (c) 2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Connection for SQL Server
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.10
'* Create Time: 18/5/2021
'* 1.0.2	18/6/2021	Modify OpenOrKeepActive
'* 1.0.3	19/6/2021	Modify OpenOrKeepActive, ConnStatusEnum,IsDBConnReady and add mIsDBOnline,RefMirrSrvTime,LastRefMirrSrvTime
'* 1.0.4	20/6/2021	Modify OpenOrKeepActive, Add mConnClose,mConnOpen
'* 1.0.5	21/6/2021	Modify mIsDBOnline
'* 1.0.6	21/7/2021	Modify mSetConnSQLServer
'* 1.1		29/8/2021   Add support for .net core
'* 1.2		24/9/2021   Add PigKeyValueApp,InitPigKeyValue
'* 1.3		5/10/2021   Modify InitPigKeyValue
'* 1.5		5/12/2021   Modify OpenOrKeepActive
'* 1.6		6/12/2021   Add IsEncrypt,OpenOrKeepActive
'* 1.7		15/12/2021	Rewrite the error handling code with LOG.
'* 1.8		28/12/2021	Increase initial value of internal variable
'* 1.9		5/1/2022	Modify InitPigKeyValue
'* 1.10		20/5/2022	Modify OpenOrKeepActive
'**********************************
Imports System.Data
Imports PigKeyCacheLib
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
#Else
Imports Microsoft.Data.SqlClient
#End If
Imports PigToolsLiteLib

Public Class ConnSQLSrv
	Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.10.3"
    Public Connection As SqlConnection
	Public PigKeyValueApp As PigKeyValueApp
	Private mcstChkDBStatus As CmdSQLSrvText

	Public Enum ConnStatusEnum
		Unknow = 0
		PrincipalOnline = 10
		MirrorOnline = 20
		Offline = 30
	End Enum

	Public Enum RunModeEnum
		Mirror = 10
		StandAlone = 20
	End Enum

	Private Property mLastConnSQLServer As String

	Private mintRunMode As RunModeEnum
	Public Property RunMode() As RunModeEnum
		Get
			Return mintRunMode
		End Get
		Friend Set(ByVal value As RunModeEnum)
			mintRunMode = value
		End Set
	End Property

	''' <summary>
	''' Time to refresh the mirror database, in seconds.
	''' </summary>
	Private mintRefMirrSrvTime As Integer = 30
	Public Property RefMirrSrvTime() As Integer
		Get
			Return mintRefMirrSrvTime
		End Get
		Set(ByVal value As Integer)
			If value <= 0 Then
				mintRefMirrSrvTime = 30
			Else
				mintRefMirrSrvTime = value
			End If
		End Set
	End Property

	''' <summary>
	''' The last time the mirror database was refreshed
	''' </summary>
	Private mdteLastRefMirrSrvTime As DateTime
	Public Property LastRefMirrSrvTime() As DateTime
		Get
			Return mdteLastRefMirrSrvTime
		End Get
		Friend Set(ByVal value As DateTime)
			mdteLastRefMirrSrvTime = value
		End Set
	End Property

	''' <summary>
	''' If Mirror SQL server is not specified, it will run in stand-alone mode.
	''' </summary>
	Private mstrPrincipalSQLServer As String = ""
	Public Property PrincipalSQLServer() As String
		Get
			Return mstrPrincipalSQLServer
		End Get
		Friend Set(ByVal value As String)
			mstrPrincipalSQLServer = value
		End Set
	End Property

	''' <summary>
	''' If Mirror SQL server is specified, it will run in mirror mode and can automatic failover.
	''' </summary>
	Private mstrMirrorSQLServer As String = ""
	Public Property MirrorSQLServer() As String
		Get
			Return mstrMirrorSQLServer
		End Get
		Friend Set(ByVal value As String)
			mstrMirrorSQLServer = value
		End Set
	End Property

	''' <summary>
	''' If running in mirror mode, the current database of the principal server and the mirror server must be the same.
	''' </summary>
	Private mstrCurrDatabase As String = ""
	Public Property CurrDatabase() As String
		Get
			Return mstrCurrDatabase
		End Get
		Friend Set(ByVal value As String)
			mstrCurrDatabase = value
		End Set
	End Property

	''' <summary>
	''' If running in mirror mode, the uid of the principal server and the mirror server must be the same.
	''' </summary>
	Private mstrDBUser As String = ""
	Public Property DBUser() As String
		Get
			Return mstrDBUser
		End Get
		Friend Set(ByVal value As String)
			mstrDBUser = value
		End Set
	End Property

	''' <summary>
	''' If running in mirror mode, the password of the principal server and the mirror server must be the same.
	''' </summary>
	Private mstrDBUserPwd As String = ""
	Public Property DBUserPwd() As String
		Get
			Return mstrDBUserPwd
		End Get
		Friend Set(ByVal value As String)
			mstrDBUserPwd = value
		End Set
	End Property

	''' <summary>
	''' Trusted Connectionst and stand-alone mode
	''' </summary>
	''' <param name="SQLServer">SQL Server hostname or ip</param>
	''' <param name="CurrDatabase">current database</param>
	Public Sub New(SQLServer As String, CurrDatabase As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(SQLServer, CurrDatabase)
	End Sub

	''' <summary>
	''' Trusted Connectionst and mirror mode
	''' </summary>
	''' <param name="PrincipalSQLServer">Principal SQLServer hostname or ip</param>
	''' <param name="MirrorSQLServer">Mirror SQLServer hostname or ip</param>
	''' <param name="CurrDatabase">current database</param>
	''' <param name="Provider">What driver to use</param>
	Public Sub New(PrincipalSQLServer As String, MirrorSQLServer As String, CurrDatabase As String)
		MyBase.New(CLS_VERSION)
		Me.MirrorSQLServer = MirrorSQLServer
		Me.mNew(PrincipalSQLServer, CurrDatabase)
	End Sub


	''' <summary>
	''' Database user password login Connectionst and stand-alone mode
	''' </summary>
	''' <param name="SQLServer">SQL Server hostname or ip</param>
	''' <param name="DBUser">Database user</param>
	''' <param name="DBUserPwd">Database user password</param>
	Public Sub New(SQLServer As String, CurrDatabase As String, DBUser As String, DBUserPwd As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(SQLServer, CurrDatabase, DBUser, DBUserPwd)
	End Sub

	''' <summary>
	''' Database user password login Connectionst and mirror mode
	''' </summary>
	''' <param name="PrincipalSQLServer">Principal SQLServer hostname or ip</param>
	''' <param name="MirrorSQLServer">Mirror SQLServer hostname or ip</param>
	''' <param name="CurrDatabase">current database</param>
	''' <param name="DBUser">Database user</param>
	''' <param name="DBUserPwd">Database user password</param>
	Public Sub New(PrincipalSQLServer As String, MirrorSQLServer As String, CurrDatabase As String, DBUser As String, DBUserPwd As String)
		MyBase.New(CLS_VERSION)
		Me.MirrorSQLServer = MirrorSQLServer
		Me.mNew(PrincipalSQLServer, CurrDatabase, DBUser, DBUserPwd)
	End Sub

	Private mbolIsTrustedConnection As Boolean
	Public Property IsTrustedConnection() As Boolean
		Get
			Return mbolIsTrustedConnection
		End Get
		Friend Set(ByVal value As Boolean)
			mbolIsTrustedConnection = value
		End Set
	End Property


	Private Sub mNew(PrincipalSQLServer As String, CurrDatabase As String, Optional DBUser As String = "", Optional DBUserPwd As String = "")
		Dim strStepName As String = ""
		Try
			With Me
				.PrincipalSQLServer = PrincipalSQLServer
				.CurrDatabase = CurrDatabase
				If DBUser = "" Then
					.IsTrustedConnection = True
				Else
					.IsTrustedConnection = False
					.DBUser = DBUser
					.DBUserPwd = DBUserPwd
				End If
				If .MirrorSQLServer = "" Then
					.RunMode = RunModeEnum.StandAlone
				Else
					.RunMode = RunModeEnum.Mirror
				End If
				.ConnectionTimeout = 5
				.CommandTimeout = 60
				Me.Connection = New SqlConnection
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mNew", strStepName, ex)
		End Try
	End Sub

	''' <summary>
	''' Database connection status, including the connection between principal and mirror server.
	''' </summary>
	Private mintConnStatus As ConnStatusEnum = ConnStatusEnum.Unknow
	Public Property ConnStatus() As ConnStatusEnum
		Get
			Return mintConnStatus
		End Get
		Friend Set(ByVal value As ConnStatusEnum)
			mintConnStatus = value
		End Set
	End Property

	''' <summary>
	''' Timeout for executing SQL command
	''' </summary>
	Private mlngCommandTimeout As Long
	Public Property CommandTimeout() As Long
		Get
			Return mlngCommandTimeout
		End Get
		Set(ByVal value As Long)
			mlngCommandTimeout = value
		End Set
	End Property

	''' <summary>
	''' Connection database timeout
	''' </summary>
	Private mlngConnectionTimeout As Long
	Public Property ConnectionTimeout() As Long
		Get
			Return mlngConnectionTimeout
		End Get
		Set(ByVal value As Long)
			mlngConnectionTimeout = value
		End Set
	End Property

	Private Sub mConnClose()
		Try
			Me.Connection.Close()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mConnClose", ex)
		End Try
	End Sub

	Private Sub mConnOpen()
		Try
			Me.Connection.Open()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mConnOpen", ex)
		End Try
	End Sub

	''' <summary>
	''' Open or keep the database connection available
	''' </summary>
	Public Sub OpenOrKeepActive()
		Dim LOG As New PigStepLog("OpenOrKeepActive")
		Try
			Select Case Me.RunMode
				Case RunModeEnum.StandAlone
					If Me.Connection Is Nothing Then
						LOG.StepName = "New SqlConnection"
						Me.Connection = New SqlConnection
					End If
					With Me.Connection
						Select Case .State
							Case ConnectionState.Closed
								LOG.StepName = "SetConnSQLServer"
								If Me.IsTrustedConnection = True Then
									Me.mSetConnSQLServer(Me.PrincipalSQLServer, Me.CurrDatabase)
								Else
									Me.mSetConnSQLServer(Me.PrincipalSQLServer, Me.DBUser, Me.DBUserPwd, Me.CurrDatabase)
								End If
								If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
								.ConnectionString &= "Connect Timeout=" & Me.ConnectionTimeout & ";"
								If Me.IsEncrypt = True Then
									.ConnectionString &= "Encrypt=True;"
								Else
									.ConnectionString &= "Encrypt=False;"
								End If
								LOG.StepName = "Open"
								Me.mConnOpen()
								If Me.LastErr <> "" Then
									Me.ConnStatus = ConnStatusEnum.Offline
									Throw New Exception(Me.LastErr)
								End If
								Me.ConnStatus = ConnStatusEnum.PrincipalOnline
						End Select
					End With
				Case RunModeEnum.Mirror
					If Me.MirrorSQLServer = "" Then Throw New Exception("Mirror SQLServer is not defined")
					Dim bolIsConn As Boolean = False
					Select Case Me.ConnStatus
						Case ConnStatusEnum.Unknow, ConnStatusEnum.Offline
							If Me.mLastConnSQLServer = "" Or mLastConnSQLServer = Me.MirrorSQLServer Then
								Me.mLastConnSQLServer = Me.PrincipalSQLServer
							Else
								Me.mLastConnSQLServer = Me.MirrorSQLServer
							End If
							bolIsConn = True
						Case Else
							If Math.Abs(DateDiff("s", Me.LastRefMirrSrvTime, Now)) > Me.RefMirrSrvTime Then
								If Me.mIsDBOnline = True Then
									Me.LastRefMirrSrvTime = Now
								Else
									If Me.ConnStatus = ConnStatusEnum.PrincipalOnline Then
										Me.mLastConnSQLServer = Me.MirrorSQLServer
									Else
										Me.mLastConnSQLServer = Me.PrincipalSQLServer
									End If
									bolIsConn = True
								End If
							End If
					End Select
					If bolIsConn = True Then
						If Not Me.Connection Is Nothing Then
							If Me.Connection.State <> ConnectionState.Closed Then
								Me.mConnClose()
							End If
							Me.Connection = Nothing
						End If
						LOG.StepName = "New SqlConnection"
						Me.Connection = New SqlConnection
						With Me.Connection
							LOG.StepName = "SetConnSQLServer first time"
							If Me.IsTrustedConnection = True Then
								Me.mSetConnSQLServer(Me.mLastConnSQLServer, Me.CurrDatabase)
							Else
								Me.mSetConnSQLServer(Me.mLastConnSQLServer, Me.DBUser, Me.DBUserPwd, Me.CurrDatabase)
							End If
							If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
							.ConnectionString &= "Connect Timeout=" & Me.ConnectionTimeout & ";"
                            If Me.IsEncrypt = True Then
                                .ConnectionString &= "Encrypt=True;"
                            Else
                                .ConnectionString &= "Encrypt=False;"
                            End If
                            LOG.StepName = "Open first time"
                            Me.mConnOpen()
							If Me.LastErr = "" Then
								If Me.mIsDBOnline = True Then
									If Me.mLastConnSQLServer = Me.PrincipalSQLServer Then
										Me.ConnStatus = ConnStatusEnum.PrincipalOnline
									Else
										Me.ConnStatus = ConnStatusEnum.MirrorOnline
									End If
									Me.LastRefMirrSrvTime = Now
								End If
								bolIsConn = False
							End If
						End With
						If bolIsConn = True Then
							If Me.mLastConnSQLServer = "" Or mLastConnSQLServer = Me.MirrorSQLServer Then
								Me.mLastConnSQLServer = Me.PrincipalSQLServer
							Else
								Me.mLastConnSQLServer = Me.MirrorSQLServer
							End If
							With Me.Connection
								LOG.StepName = "SetConnSQLServer second time"
								If Me.IsTrustedConnection = True Then
									Me.mSetConnSQLServer(Me.mLastConnSQLServer, Me.CurrDatabase)
								Else
									Me.mSetConnSQLServer(Me.mLastConnSQLServer, Me.DBUser, Me.DBUserPwd, Me.CurrDatabase)
								End If
								If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
								.ConnectionString &= "Connect Timeout=" & Me.ConnectionTimeout & ";"
								LOG.StepName = "Open second time"
								Me.mConnOpen()
								If Me.LastErr = "" Then
									LOG.StepName = "mIsDBOnline second time"
									If Me.mIsDBOnline = True Then
										If Me.mLastConnSQLServer = Me.PrincipalSQLServer Then
											Me.ConnStatus = ConnStatusEnum.PrincipalOnline
										Else
											Me.ConnStatus = ConnStatusEnum.MirrorOnline
										End If
										Me.LastRefMirrSrvTime = Now
									Else
										Me.ConnStatus = ConnStatusEnum.Offline
										Throw New Exception(Me.LastErr)
									End If
								Else
									Me.ConnStatus = ConnStatusEnum.Offline
									Throw New Exception(Me.LastErr)
								End If
							End With
						End If
					End If
				Case Else
					Throw New Exception("Unknow run mode")
			End Select
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			If Me.ConnStatus <> ConnStatusEnum.Offline Then Me.ConnStatus = ConnStatusEnum.Unknow
		End Try
	End Sub

	Private mbolIsEncrypt As Boolean = False
	Public Property IsEncrypt() As Boolean
		Get
			Return mbolIsEncrypt
		End Get
		Set(value As Boolean)
			mbolIsEncrypt = value
		End Set
	End Property

	Public ReadOnly Property IsDBConnReady() As Boolean
		Get
			Try
				Select Case Me.ConnStatus
					Case ConnStatusEnum.PrincipalOnline, ConnStatusEnum.MirrorOnline
						Return True
					Case Else
						Return False
				End Select
			Catch ex As Exception
				Me.SetSubErrInf("IsDBConnReady", ex)
				Return False
			End Try
		End Get
	End Property

	Private Sub mSetConnSQLServer(SQLServer As String, DBUser As String, DBUserPwd As String, CurrDatabase As String)
		Try
			Dim strConn As String = ""
			DBUserPwd = Replace(DBUserPwd, "'", "''")
			strConn = "Server=" & SQLServer & ";Database=" & CurrDatabase & ";Uid='" & DBUser & "';Pwd='" & DBUserPwd & "';"
			strConn &= "MultipleActiveResultSets=true;"
			Me.Connection.ConnectionString = strConn
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mSetConnSQLServer", ex)
		End Try
	End Sub

	Private Sub mSetConnSQLServer(SQLServer As String, CurrDatabase As String)
		Try
			Dim strConn As String = "Server=" & SQLServer & ";Database=" & CurrDatabase & ";Trusted_Connection=true;"
			strConn &= "MultipleActiveResultSets=true;"
			Me.Connection.ConnectionString = strConn
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mSetConnSQLServer", ex)
		End Try
	End Sub

	Private Function mIsDBOnline() As Boolean
		Dim LOG As New PigStepLog("mIsDBOnline")
		Try
			If Me.Connection Is Nothing Then Throw New Exception("No connection established")
			If Me.Connection.State <> ConnectionState.Open Then Throw New Exception("The current connection status is " & Me.Connection.State.ToString)
			If mcstChkDBStatus Is Nothing Then
				LOG.StepName = "New CmdSQLSrvText"
				mcstChkDBStatus = New CmdSQLSrvText("SELECT Convert(varchar(50),DatabasePropertyEx(@DBName,'status')) DBStatus")
				If mcstChkDBStatus.LastErr <> "" Then Throw New Exception(mcstChkDBStatus.LastErr)
				LOG.StepName = "AddPara(@DBName)"
				mcstChkDBStatus.AddPara("@DBName", SqlDbType.NVarChar, 512)
				If mcstChkDBStatus.LastErr <> "" Then Throw New Exception(mcstChkDBStatus.LastErr)
				LOG.StepName = "Set ActiveConnection"
				mcstChkDBStatus.ActiveConnection = Me.Connection
				If mcstChkDBStatus.LastErr <> "" Then Throw New Exception(mcstChkDBStatus.LastErr)
			End If
			Dim rsAny As Recordset
			mcstChkDBStatus.ParaValue("@DBName") = Me.CurrDatabase
			LOG.StepName = "Execute"
			rsAny = mcstChkDBStatus.Execute
			If mcstChkDBStatus.LastErr <> "" Then Throw New Exception(mcstChkDBStatus.LastErr)
			Dim strDBStaus As String = UCase(rsAny.Fields.Item("DBStatus").StrValue)
			rsAny.Close()
			rsAny = Nothing
			If strDBStaus = "ONLINE" Then
				Return True
			Else
				Return False
			End If
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return False
		End Try
	End Function

	Public Sub InitPigKeyValue()
		Try
			If Me.IsWindows = True Then
				Me.PigKeyValueApp = New PigKeyValueApp(Me.Connection.ConnectionString, PigKeyValueApp.EnmCacheLevel.ToShareMem)
			Else
				Me.PigKeyValueApp = New PigKeyValueApp()
			End If
			If Me.PigKeyValueApp.LastErr <> "" Then Throw New Exception(Me.PigKeyValueApp.LastErr)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("InitPigKeyValue", ex)
		End Try
	End Sub

End Class
