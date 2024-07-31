'**********************************
'* Name: DBConnMgr
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Database connection management
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.12
'* Create Time: 17/10/2021
'* 1.1	1/2/2022	Modify New
'* 1.2	23/3/2022	Modify New, add MkEncKey,LoadDBConnDefs
'* 1.3	10/4/2022	Add SaveDBConnDefs
'* 1.4	11/4/2022	Modify New
'* 1.5	12/4/2022	Modify New
'* 1.6	1/5/2022	Add IsConfigChange, modify New,fPigConfigApp
'* 1.7	20/5/2022	Modify SaveDBConnDefs,LoadDBConnDefs
'* 1.8	22/5/2022	Modify SaveDBConnDefs,LoadDBConnDefs
'* 1.9	2/7/2022	Use PigBaseLocal
'* 1.10	26/7/2022	Modify Imports
'* 1.11	29/7/2022	Modify Imports
'* 1.12 28/7/2024   Modify PigStepLog to StruStepLog
'**********************************
Imports PigToolsLiteLib

Friend Class DBConnMgr
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1." & "12" & "." & "2"
    Friend Property fPigConfigApp As PigConfigApp
    Private Property mConfFilePath As String
    Public ReadOnly Property DBConnDefs As DBConnDefs

    Public Sub New(ConfFilePath As String)
        MyBase.New(CLS_VERSION)
        Me.mConfFilePath = ConfFilePath
        Me.fPigConfigApp = New PigConfigApp(PigText.enmTextType.UTF8)
        'Me.DBConnDefs = New DBConnDefs(Me)
    End Sub

    Public Sub New(EncKey As String, ConfFilePath As String)
        MyBase.New(CLS_VERSION)
        Me.mConfFilePath = ConfFilePath
        Me.fPigConfigApp = New PigConfigApp(EncKey, PigText.enmTextType.UTF8)
        'Me.DBConnDefs = New DBConnDefs(Me)
    End Sub

    Public Function MkEncKey(ByRef Base64EncKey As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "MkEncKey"
        Try
            LOG.StepName = "MkEncKey"
            LOG.Ret = Me.fPigConfigApp.MkEncKey(Base64EncKey)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "ReNew PigConfigApp"
            Me.fPigConfigApp = New PigConfigApp(Base64EncKey, PigText.enmTextType.UTF8)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function



    Public Function LoadDBConnDefs() As String
        Dim LOG As New StruStepLog : LOG.SubName = "LoadDBConnDefs"
        Try
            LOG.StepName = "LoadConfigFile"
            LOG.Ret = Me.fPigConfigApp.LoadConfigFile(Me.mConfFilePath, PigConfigApp.EnmSaveType.Xml)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(Me.mConfFilePath)
                Throw New Exception(LOG.Ret)
            End If
            LoadDBConnDefs = ""
            For Each oPigConfigSession As PigConfigSession In Me.fPigConfigApp.PigConfigSessions
                With oPigConfigSession
                    If .SessionName <> "Main" Then
                        Dim strPrincipalSQLServer As String = ""
                        If .PigConfigs.IsItemExists("PrincipalSQLServer") = True Then
                            strPrincipalSQLServer = .PigConfigs.Item("PrincipalSQLServer").ConfValue
                        ElseIf .PigConfigs.IsItemExists("PrincipalSQLServer") = True Then
                            strPrincipalSQLServer = .PigConfigs.Item("SQLServer").ConfValue
                        End If
                        If strPrincipalSQLServer = "" Then
                            LoadDBConnDefs &= .SessionName & ".PrincipalSQLServer Undefined;"
                        Else
                            Dim strCurrDatabase As String = ""
                            If .PigConfigs.IsItemExists("CurrDatabase") = True Then strCurrDatabase = .PigConfigs.Item("CurrDatabase").ConfValue
                            If strCurrDatabase = "" Then strCurrDatabase = "master"
                            Dim intRunMode As ConnSQLSrv.RunModeEnum
                            If .PigConfigs.IsItemExists("MirrorSQLServer") = False Then
                                intRunMode = ConnSQLSrv.RunModeEnum.StandAlone
                            ElseIf .PigConfigs.Item("MirrorSQLServer").ConfValue = "" Then
                                intRunMode = ConnSQLSrv.RunModeEnum.StandAlone
                            Else
                                intRunMode = ConnSQLSrv.RunModeEnum.Mirror
                            End If
                            Dim bolIsTrustedConnection As Boolean
                            If .PigConfigs.IsItemExists("DBUser") = False Then
                                bolIsTrustedConnection = True
                            ElseIf .PigConfigs.Item("DBUser").ConfValue = "" Then
                                bolIsTrustedConnection = True
                            Else
                                bolIsTrustedConnection = False
                            End If
                            If Me.DBConnDefs.IsItemExists(.SessionName) = True Then
                                LOG.StepName = "DBConnDefs.Remove"
                                LOG.Ret = Me.DBConnDefs.Remove(.SessionName)
                                If LOG.Ret <> "OK" Then
                                    LOG.AddStepNameInf(.SessionName)
                                    LoadDBConnDefs &= LOG.StepName & .SessionName & LOG.Ret & ";"
                                End If
                            End If
                            Select Case intRunMode
                                Case ConnSQLSrv.RunModeEnum.StandAlone
                                    If bolIsTrustedConnection = True Then
                                        Me.DBConnDefs.Add(.SessionName, strPrincipalSQLServer, strCurrDatabase, .SessionDesc)
                                    Else
                                        Me.DBConnDefs.Add(.SessionName, strPrincipalSQLServer, strCurrDatabase, .PigConfigs.Item("DBUser").ConfValue, .PigConfigs.Item("DBUserPwd").ConfValue, .SessionDesc)
                                    End If
                                Case ConnSQLSrv.RunModeEnum.Mirror
                                    If bolIsTrustedConnection = True Then
                                        Me.DBConnDefs.Add(.SessionName, strPrincipalSQLServer, .PigConfigs.Item("MirrorSQLServer").ConfValue, strCurrDatabase, .SessionDesc)
                                    Else
                                        Me.DBConnDefs.Add(.SessionName, strPrincipalSQLServer, .PigConfigs.Item("MirrorSQLServer").ConfValue, strCurrDatabase, .PigConfigs.Item("DBUser").ConfValue, .PigConfigs.Item("DBUserPwd").ConfValue, .SessionDesc)
                                    End If
                            End Select
                        End If
                    End If
                End With
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function SaveDBConnDefs() As String
        Dim LOG As New StruStepLog : LOG.SubName = "SaveDBConnDefs"
        Try

            LOG.StepName = "LoadConfigFile"
            LOG.Ret = Me.fPigConfigApp.LoadConfigFile(Me.mConfFilePath, PigConfigApp.EnmSaveType.Xml)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(Me.mConfFilePath)
                Throw New Exception(LOG.Ret)
            End If
            For Each oPigConfigSession As PigConfigSession In Me.fPigConfigApp.PigConfigSessions
                With oPigConfigSession
                    If .SessionName = "Main" Then

                    Else

                    End If
                End With
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
