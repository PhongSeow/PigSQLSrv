'**********************************
'* Name: DBConnDef
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Database connection definition
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.8
'* Create Time: 17/10/2021
'* 1.1	1/2/2022	Modify New, add Properties
'* 1.2	1/2/2022	Add IsTrustedConnection, mNew
'* 1.3	10/4/2022	Add New, modify RunMode
'* 1.4	20/5/2022	Add DBConnDesc
'* 1.5	22/5/2022	Modify New, add fPigConfigSession
'* 1.6	2/7/2022	Use PigBaseLocal
'* 1.7	26/7/2022	Modify Imports
'* 1.8	29/7/2022	Modify Imports
'**********************************
#If NETFRAMEWORK Then
Imports PigToolsWinLib
#Else
Imports PigToolsLiteLib
#End If


Friend Class DBConnDef
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.8.1"

    Friend fPigConfigSession As PigConfigSession

    Public ReadOnly Property RunMode As ConnSQLSrv.RunModeEnum
        Get
            Try
                If Me.fPigConfigSession.PigConfigs.Item("MirrorSQLServer").ConfValue = "" Then
                    Return ConnSQLSrv.RunModeEnum.StandAlone
                Else
                    Return ConnSQLSrv.RunModeEnum.Mirror
                End If
            Catch ex As Exception
                Me.SetSubErrInf("RunMode.Get", ex)
                Return ConnSQLSrv.RunModeEnum.StandAlone
            End Try
        End Get
    End Property

    Public ReadOnly Property DBConnName As String
        Get
            Try
                Return Me.fPigConfigSession.SessionName
            Catch ex As Exception
                Me.SetSubErrInf("DBConnName.Get", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property DBConnDesc As String
        Get
            Try
                Return Me.fPigConfigSession.SessionDesc
            Catch ex As Exception
                Me.SetSubErrInf("DBConnDesc.Get", ex)
                Return ""
            End Try
        End Get
    End Property

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public Property CurrDatabase As String
        Get
            Try
                Return Me.fPigConfigSession.PigConfigs.Item("CurrDatabase").ConfValue
            Catch ex As Exception
                Me.SetSubErrInf("CurrDatabase.Get", ex)
                Return ""
            End Try
        End Get
        Set(value As String)
            Try
                Me.fPigConfigSession.PigConfigs.Item("CurrDatabase").ConfValue = value
            Catch ex As Exception
                Me.SetSubErrInf("CurrDatabase.Set", ex)
            End Try
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
