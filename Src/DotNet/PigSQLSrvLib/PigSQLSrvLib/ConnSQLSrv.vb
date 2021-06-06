'**********************************
'* Name: ConnSQLSrv
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Connection for SQL Server
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Create Time: 18/5/2021
'**********************************
Imports System.Data
Imports System.Data.SqlClient

Public Class ConnSQLSrv
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.1"
	Public Connection As SqlConnection


	'Public Enum SQLSrvProviderEnum
	'	MicrosoftSQLServer = 90
	'	MicrosoftSQLServer2012NativeClient = 100
	'End Enum

	Public Enum RunModeEnum
		Mirror = 10
		StandAlone = 20
	End Enum

	Public Enum ConnStatusEnum
		Unknow = 0
		PrincipalOnline = 10
		PrincipalAndMirrorOnline = 20
		PrincipalOnlineMirrorOffline = 30
		PrincipalOfflineMirrorOnline = 40
		PrincipalAndMirrorOffline = 50
	End Enum

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
	''' If Mirror SQL server is not specified, it will run in stand-alone mode.
	''' </summary>
	Private mstrPrincipalSQLServer As String
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
	Private mstrMirrorSQLServer As String
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
	Private mstrCurrDatabase As String
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
	Private mstrDBUser As String
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
	Private mstrDBUserPwd As String
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
	''' Database user password login Connectionst and stand-alone mode
	''' </summary>
	''' <param name="SQLServer">SQL Server hostname or ip</param>
	''' <param name="CurrDatabase">current database</param>
	Public Sub New(SQLServer As String, CurrDatabase As String, DBUser As String, DBUserPwd As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(SQLServer, CurrDatabase, DBUser, DBUserPwd)
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

	''' <summary>
	''' Open or keep the database connection available
	''' </summary>
	Public Sub OpenOrKeepActive()
		Dim strStepName As String = ""
		Try
			With Me.Connection
				Select Case Me.RunMode
					Case RunModeEnum.StandAlone
						Select Case .State
							Case ConnectionState.Closed
								strStepName = "SetConnSQLServer"
								If Me.IsTrustedConnection = True Then
									Me.mSetConnSQLServer(Me.PrincipalSQLServer, Me.CurrDatabase)
								Else
									Me.mSetConnSQLServer(Me.PrincipalSQLServer, Me.DBUser, Me.DBUserPwd, Me.CurrDatabase)
								End If
								If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
								.ConnectionString &= "Connect Timeout=" & Me.ConnectionTimeout & ";"
								strStepName = "Open"
								.Open()
								Me.ConnStatus = ConnStatusEnum.PrincipalOnline
						End Select
					Case RunModeEnum.Mirror
						Throw New Exception("Not support now")
					Case Else
						Throw New Exception("Unknow run mode")
				End Select
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("OpenOrKeepActive", strStepName, ex)
			Me.ConnStatus = ConnStatusEnum.Unknow
		End Try
	End Sub

	Public ReadOnly Property IsDBConnReady() As Boolean
		Get
			Try
				Select Case Me.ConnStatus
					Case ConnStatusEnum.PrincipalAndMirrorOnline, ConnStatusEnum.PrincipalOfflineMirrorOnline, ConnStatusEnum.PrincipalOnline, ConnStatusEnum.PrincipalOnlineMirrorOffline
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
			Me.Connection.ConnectionString = strConn
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mSetConnSQLServer", ex)
		End Try
	End Sub

	Private Sub mSetConnSQLServer(SQLServer As String, CurrDatabase As String)
		Try
			Dim strConn As String = "Server=" & SQLServer & ";Database=" & CurrDatabase & ";Trusted_Connection=true;"
			Me.Connection.ConnectionString = strConn
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mSetConnSQLServer", ex)
		End Try
	End Sub

End Class
