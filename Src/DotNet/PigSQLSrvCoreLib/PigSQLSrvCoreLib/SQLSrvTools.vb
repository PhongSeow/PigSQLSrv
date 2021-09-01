'**********************************
'* Name: SQLSrvTools
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Common SQL server tools
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 1/9/2021
'* 1.0		1/9/2021   Add IsDBObjExists,IsDBUserExists,IsDatabaseExists,IsLoginUserExists
'**********************************
Imports System.Data
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
#Else
Imports Microsoft.Data.SqlClient
#End If
Public Class SQLSrvTools
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.5"
    Private moConnSQLSrv As ConnSQLSrv

    Public Enum enmDBObjType
        Unknow = 0
        UserTable = 10
        View = 20
        StoredProcedure = 30
        ScalarFunction = 40
        InlineFunction = 50
    End Enum

    Public Sub New(ConnSQLSrv As ConnSQLSrv)
        MyBase.New(CLS_VERSION)
        Try
            moConnSQLSrv = ConnSQLSrv
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    Public Function IsDBObjExists(DBObjType As enmDBObjType, ObjName As String) As Boolean
        Const SUB_NAME As String = "IsDBObjExists"
        Dim strStepName As String = ""
        Try
            Dim strXType As String = ""
            strStepName = "Check DBObjType"
            Select Case DBObjType
                Case enmDBObjType.UserTable
                    strXType = "U"
                Case enmDBObjType.View
                    strXType = "V"
                Case enmDBObjType.StoredProcedure
                    strXType = "P"
                Case enmDBObjType.ScalarFunction
                    strXType = "FN"
                Case enmDBObjType.InlineFunction
                    strXType = "IF"
                Case Else
                    Throw New Exception("Cannot support")
            End Select
            Dim strSQL As String = "select 1 from sysobjects where name=@ObjName and xtype=@DBObjType"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@ObjName", SqlDbType.VarChar, 512)
                .AddPara("@DBObjType", SqlDbType.VarChar, 10)
                .ParaValue("@ObjName") = ObjName
                .ParaValue("@DBObjType") = strXType
                Dim rsAny = .Execute()
                If .LastErr <> "" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(.LastErr)
                End If
                If rsAny.EOF = True Then
                    IsDBObjExists = False
                Else
                    IsDBObjExists = True
                End If
            End With
            oCmdSQLSrvText = Nothing
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return False
        End Try
    End Function


    Public Function IsDatabaseExists(DBName As String) As Boolean
        Const SUB_NAME As String = "IsDatabaseExists"
        Dim strStepName As String = ""
        Try
            Dim strSQL As String = "select 1 from master.dbo.sysdatabases where name=@DBName"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@DBName", SqlDbType.VarChar, 512)
                .ParaValue("@DBName") = DBName
                Dim rsAny = .Execute()
                If .LastErr <> "" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(.LastErr)
                End If
                If rsAny.EOF = True Then
                    Return False
                Else
                    Return True
                End If
            End With
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return False
        End Try
    End Function

    Public Function IsLoginUserExists(LoginName As String) As Boolean
        Const SUB_NAME As String = "IsLoginUserExists"
        Dim strStepName As String = ""
        Try
            Dim strSQL As String = "select 1 from master.dbo.syslogins where name=@LoginName"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@LoginName", SqlDbType.VarChar, 512)
                .ParaValue("@LoginName") = LoginName
                Dim rsAny = .Execute()
                If .LastErr <> "" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(.LastErr)
                End If
                If rsAny.EOF = True Then
                    Return False
                Else
                    Return True
                End If
            End With
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return False
        End Try
    End Function

    Public Function IsDBUserExists(DBUserName As String) As Boolean
        Const SUB_NAME As String = "IsDBUserExists"
        Dim strStepName As String = ""
        Try
            Dim strSQL As String = "select 1 from sysusers where name=@DBUserName and islogin=1"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@DBUserName", SqlDbType.VarChar, 512)
                .ParaValue("@DBUserName") = DBUserName
                Dim rsAny = .Execute()
                If .LastErr <> "" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(.LastErr)
                End If
                If rsAny.EOF = True Then
                    Return False
                Else
                    Return True
                End If
            End With
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return False
        End Try
    End Function

End Class
