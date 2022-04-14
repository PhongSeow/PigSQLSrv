'**********************************
'* Name: DBConnMgr
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Database connection management
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 17/10/2021
'* 1.1	1/2/2022	Modify New
'* 1.2	23/3/2022	Modify New, add MkEncKey,LoadDBConnDefs
'* 1.3	10/4/2022	Add SaveDBConnDefs
'* 1.4	11/4/2022	Modify New
'* 1.5	12/4/2022	Modify New
'**********************************
Imports PigToolsLiteLib
Public Class DBConnMgr
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.5.5"
    Private Property mPigConfigApp As PigConfigApp
    Private Property mConfFilePath As String
    Public ReadOnly Property DBConnDefs As DBConnDefs

    Public Sub New(ConfFilePath As String)
        MyBase.New(CLS_VERSION)
        Me.mConfFilePath = ConfFilePath
        Me.mPigConfigApp = New PigConfigApp(PigText.enmTextType.UTF8)
        Me.DBConnDefs = New DBConnDefs
    End Sub

    Public Sub New(EncKey As String, ConfFilePath As String)
        MyBase.New(CLS_VERSION)
        Me.mConfFilePath = ConfFilePath
        Me.mPigConfigApp = New PigConfigApp(EncKey, PigText.enmTextType.UTF8)
        Me.DBConnDefs = New DBConnDefs
    End Sub

    Public Function MkEncKey(ByRef Base64EncKey As String) As String
        Dim LOG As New PigStepLog("MkEncKey")
        Try
            LOG.StepName = "MkEncKey"
            LOG.Ret = Me.mPigConfigApp.MkEncKey(Base64EncKey)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "ReNew PigConfigApp"
            Me.mPigConfigApp = New PigConfigApp(Base64EncKey, PigText.enmTextType.UTF8)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Public Function LoadDBConnDefs() As String
		Dim LOG As New PigStepLog("LoadDBConnDefs")
        Try
            LOG.StepName = "LoadConfigFile"
            LOG.Ret = Me.mPigConfigApp.LoadConfigFile(Me.mConfFilePath, PigConfigApp.EnmSaveType.Xml)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(Me.mConfFilePath)
                Throw New Exception(LOG.Ret)
            End If
            For Each oPigConfigSession As PigConfigSession In Me.mPigConfigApp.PigConfigSessions
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

    Public Function SaveDBConnDefs() As String
        Dim LOG As New PigStepLog("SaveDBConnDefs")
        Try
            LOG.StepName = "LoadConfigFile"
            LOG.Ret = Me.mPigConfigApp.LoadConfigFile(Me.mConfFilePath, PigConfigApp.EnmSaveType.Xml)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(Me.mConfFilePath)
                Throw New Exception(LOG.Ret)
            End If
            For Each oPigConfigSession As PigConfigSession In Me.mPigConfigApp.PigConfigSessions
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
