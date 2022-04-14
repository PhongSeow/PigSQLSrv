'**********************************
'* Name: DBConnDef
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Database connection definition
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3
'* Create Time: 17/10/2021
'* 1.1	1/2/2022	Modify New, add Properties
'* 1.2	1/2/2022	Add IsTrustedConnection, mNew
'* 1.3	10/4/2022	Add New, modify RunMode
'**********************************
Imports PigToolsLiteLib

Public Class DBConnDef
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.3.2"
    Private mintRunMode As ConnSQLSrv.RunModeEnum
    Public Property RunMode As ConnSQLSrv.RunModeEnum
        Get
            Return mintRunMode
        End Get
        Friend Set(value As ConnSQLSrv.RunModeEnum)
            mintRunMode = value
        End Set
    End Property

    Public ReadOnly Property DBConnName As String

    Public Sub New(DBConnName As String, SQLServer As String, CurrDatabase As String)
        MyBase.New(CLS_VERSION)
        Me.DBConnName = DBConnName
        Me.mNew(ConnSQLSrv.RunModeEnum.StandAlone, SQLServer, CurrDatabase)
    End Sub

    Public Sub New(DBConnName As String, SQLServer As String, CurrDatabase As String, DBUser As String, DBUserPwd As String)
        MyBase.New(CLS_VERSION)
        Me.DBConnName = DBConnName
        Me.mNew(ConnSQLSrv.RunModeEnum.StandAlone, SQLServer, CurrDatabase, DBUser, DBUserPwd)
    End Sub

    Public Sub New(DBConnName As String, PrincipalSQLServer As String, MirrorSQLServer As String, CurrDatabase As String)
        MyBase.New(CLS_VERSION)
        Me.DBConnName = DBConnName
        Me.mNew(ConnSQLSrv.RunModeEnum.Mirror, PrincipalSQLServer, CurrDatabase,,, MirrorSQLServer)
    End Sub

    Public Sub New(DBConnName As String, PrincipalSQLServer As String, MirrorSQLServer As String, CurrDatabase As String, DBUser As String, DBUserPwd As String)
        MyBase.New(CLS_VERSION)
        Me.DBConnName = DBConnName
        Me.mNew(ConnSQLSrv.RunModeEnum.Mirror, PrincipalSQLServer, CurrDatabase, DBUser, DBUserPwd, MirrorSQLServer)
    End Sub

    Private Sub mNew(RunMode As ConnSQLSrv.RunModeEnum, PrincipalSQLServer As String, CurrDatabase As String, Optional DBUser As String = "", Optional DBUserPwd As String = "", Optional MirrorSQLServer As String = "")
        Try
            With Me
                .RunMode = RunMode
                .PrincipalSQLServer = PrincipalSQLServer
                Select Case .RunMode
                    Case ConnSQLSrv.RunModeEnum.Mirror
                        .MirrorSQLServer = MirrorSQLServer
                    Case ConnSQLSrv.RunModeEnum.StandAlone
                    Case Else
                        Throw New Exception("Invalid ModeEnum")
                End Select
                .DBUser = DBUser
                .DBUserPwd = DBUserPwd
                .CurrDatabase = CurrDatabase
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mNew", ex)
        End Try
    End Sub


    Private mstrCurrDatabase As String
    Public Property CurrDatabase As String
        Get
            Return mstrCurrDatabase
        End Get
        Friend Set(value As String)
            mstrCurrDatabase = value
        End Set
    End Property

    Public ReadOnly Property IsTrustedConnection As Boolean
        Get
            If Me.DBUser = "" Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property


    Private mstrPrincipalSQLServer As String
    Public Property PrincipalSQLServer As String
        Get
            Return mstrPrincipalSQLServer
        End Get
        Friend Set(value As String)
            mstrPrincipalSQLServer = value
        End Set
    End Property

    Private mstrMirrorSQLServer As String
    Public Property MirrorSQLServer As String
        Get
            Return mstrMirrorSQLServer
        End Get
        Friend Set(value As String)
            mstrMirrorSQLServer = value
        End Set
    End Property


    Private mstrDBUser As String
    Public Property DBUser As String
        Get
            Return mstrDBUser
        End Get
        Friend Set(value As String)
            mstrDBUser = value
        End Set
    End Property


    Private mstrDBUserPwd As String
    Public Property DBUserPwd As String
        Get
            Return mstrDBUserPwd
        End Get
        Friend Set(value As String)
            mstrDBUserPwd = value
        End Set
    End Property

    Private mstrConnectionTimeout As Integer
    Public Property ConnectionTimeout As Integer
        Get
            Return mstrConnectionTimeout
        End Get
        Friend Set(value As Integer)
            mstrConnectionTimeout = value
        End Set
    End Property

    Private mstrCommandTimeout As Integer
    Public Property CommandTimeout As Integer
        Get
            Return mstrCommandTimeout
        End Get
        Friend Set(value As Integer)
            mstrCommandTimeout = value
        End Set
    End Property


End Class
