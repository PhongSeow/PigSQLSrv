'**********************************
'* Name: DBConnDefs
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: DBConnDef 的集合类|Collection class of DBConnDef
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.9
'* Create Time: 1/2/2022
'* 1.1	16/3/2022	Modify Add,AddOrGet
'* 1.2	16/3/2022	Modify Add,AddOrGet
'* 1.3	12/4/2022	Add AddOrGet,Add, modify New
'* 1.4	30/4/2022	Modifiy AddOrGet
'* 1.5	1/5/2022	Add Parent,IsChange, modify mAdd
'* 1.6	20/5/2022	Modify Add,AddOrGet
'* 1.7	22/5/2022	Remove AddOrGet, Add Add
'* 1.8	8/6/2022	Modify IsItemExists
'* 1.9	2/7/2022	Use PigBaseLocal
'************************************
Imports PigToolsLiteLib

Friend Class DBConnDefs
    Inherits PigBaseLocal
    Implements IEnumerable(Of DBConnDef)
    Private Const CLS_VERSION As String = "1.9.2"
    Private ReadOnly moList As New List(Of DBConnDef)

    Friend fPigConfigApp As PigConfigApp
    Public Parent As DBConnMgr


    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Try
                Return moList.Count
            Catch ex As Exception
                Me.SetSubErrInf("Count", ex)
                Return -1
            End Try
        End Get
    End Property
    Public Function GetEnumerator() As IEnumerator(Of DBConnDef) Implements IEnumerable(Of DBConnDef).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As DBConnDef
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(DBConnName As String) As DBConnDef
        Get
            Try
                Item = Nothing
                For Each oDBConnDef As DBConnDef In moList
                    If oDBConnDef.DBConnName = DBConnName Then
                        Item = oDBConnDef
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.DBConnName", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function IsItemExists(DBConnName As String) As Boolean
        Try
            IsItemExists = False
            For Each oDBConnDef As DBConnDef In moList
                If oDBConnDef.DBConnName = DBConnName Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

    Private Function mAdd(NewItem As DBConnDef) As String
        Dim LOG As New PigStepLog("mAdd")
        Try
            If Me.IsItemExists(NewItem.DBConnName) = True Then Throw New Exception(NewItem.DBConnName & " already exists.")
            LOG.StepName = "List.Add"
            moList.Add(NewItem)
            LOG.StepName = "PigConfigSessions.AddOrGet"
            Dim oPigConfigSession As PigConfigSession = Me.Parent.fPigConfigApp.PigConfigSessions.AddOrGet(NewItem.DBConnName)
            If oPigConfigSession Is Nothing Then Throw New Exception("oPigConfigSession Is Nothing")
            LOG.StepName = "PigConfigs.AddOrGet"
            With oPigConfigSession.PigConfigs
                .AddOrGet("PrincipalSQLServer", NewItem.PrincipalSQLServer)
                If .LastErr <> "" Then
                    LOG.AddStepNameInf("PrincipalSQLServer")
                    Throw New Exception(.LastErr)
                End If
                .AddOrGet("MirrorSQLServer", NewItem.MirrorSQLServer)
                If .LastErr <> "" Then
                    LOG.AddStepNameInf("MirrorSQLServer")
                    Throw New Exception(.LastErr)
                End If
                .AddOrGet("DBUser", NewItem.DBUser)
                If .LastErr <> "" Then
                    LOG.AddStepNameInf("DBUser")
                    Throw New Exception(.LastErr)
                End If
                .AddOrGet("DBUserPwd", Me.Parent.fPigConfigApp.GetEncStr(NewItem.DBUserPwd))
                If .LastErr <> "" Then
                    LOG.AddStepNameInf("DBUserPwd")
                    Throw New Exception(.LastErr)
                End If
                .AddOrGet("ConnectionTimeout", NewItem.ConnectionTimeout.ToString)
                If .LastErr <> "" Then
                    LOG.AddStepNameInf("ConnectionTimeout")
                    Throw New Exception(.LastErr)
                End If
                .AddOrGet("CommandTimeout", NewItem.CommandTimeout.ToString)
                If .LastErr <> "" Then
                    LOG.AddStepNameInf("CommandTimeout")
                    Throw New Exception(.LastErr)
                End If
            End With
            oPigConfigSession = Nothing
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(NewItem.DBConnName)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function Add(DBConnName As String, PrincipalSQLServer As String, MirrorSQLServer As String, CurrDatabase As String, DBUser As String, DBUserPwd As String, Optional DBConnDesc As String = "") As DBConnDef
        Dim LOG As New PigStepLog("Add")
        Try
            'LOG.StepName = "New DBConnDef"
            'Add = New DBConnDef(DBConnName, PrincipalSQLServer， MirrorSQLServer, CurrDatabase, DBUser, DBUserPwd)
            'If Add.LastErr <> "" Then
            '    LOG.AddStepNameInf(DBConnName)
            '    Throw New Exception(Add.LastErr)
            'End If
            'Add.DBConnDesc = DBConnDesc
            'LOG.StepName = "mAdd"
            'LOG.Ret = Me.mAdd(Add)
            'If LOG.Ret <> "OK" Then
            '    LOG.AddStepNameInf(DBConnName)
            '    Throw New Exception(LOG.Ret)
            'End If
            Me.ClearErr()
            Return Nothing
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function


    Public Function Add(DBConnName As String, SQLServer As String, CurrDatabase As String, DBUser As String, DBUserPwd As String, Optional DBConnDesc As String = "") As DBConnDef
        Dim LOG As New PigStepLog("Add")
        Try
            'LOG.StepName = "New DBConnDef"
            'Add = New DBConnDef(DBConnName, SQLServer， CurrDatabase, DBUser, DBUserPwd)
            'If Add.LastErr <> "" Then
            '    LOG.AddStepNameInf(DBConnName)
            '    Throw New Exception(Add.LastErr)
            'End If
            'Add.DBConnDesc = DBConnDesc
            'LOG.StepName = "mAdd"
            'LOG.Ret = Me.mAdd(Add)
            'If LOG.Ret <> "OK" Then
            '    LOG.AddStepNameInf(DBConnName)
            '    Throw New Exception(LOG.Ret)
            'End If
            Me.ClearErr()
            Return Nothing
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function Add(DBConnName As String, SQLServer As String, CurrDatabase As String, Optional DBConnDesc As String = "") As DBConnDef
        Dim LOG As New PigStepLog("Add")
        Try
            If Me.fPigConfigApp.PigConfigSessions.IsItemExists(DBConnName) = True Then
                Throw New Exception(DBConnName & " already exists.")
            End If
            With Me.fPigConfigApp.PigConfigSessions.Item(DBConnName)
                If .PigConfigs.IsItemExists("MirrorSQLServer") = True Then
                    LOG.StepName = "Remove(MirrorSQLServer)"
                    LOG.Ret = .PigConfigs.Remove("MirrorSQLServer")
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                End If
                .PigConfigs.Item("PrincipalSQLServer").ConfValue = SQLServer
                .PigConfigs.Item("CurrDatabase").ConfValue = CurrDatabase
                If .PigConfigs.IsItemExists("DBUser") = True Then
                    LOG.StepName = "Remove(DBUser)"
                    LOG.Ret = .PigConfigs.Remove("DBUser")
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                End If
                If .PigConfigs.IsItemExists("DBUserPwd") = True Then
                    LOG.StepName = "Remove(DBUserPwd)"
                    LOG.Ret = .PigConfigs.Remove("DBUserPwd")
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                End If
            End With
            Me.ClearErr()
            Return Nothing
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function Add(DBConnName As String, PrincipalSQLServer As String, MirrorSQLServer As String, CurrDatabase As String, Optional DBConnDesc As String = "") As DBConnDef
        Dim LOG As New PigStepLog("Add")
        Try
            'LOG.StepName = "New DBConnDef"
            'Add = New DBConnDef(DBConnName, PrincipalSQLServer， MirrorSQLServer, CurrDatabase)
            'If Add.LastErr <> "" Then
            '    LOG.AddStepNameInf(DBConnName)
            '    Throw New Exception(Add.LastErr)
            'End If
            'Add.DBConnDesc = DBConnDesc
            'LOG.StepName = "mAdd"
            'LOG.Ret = Me.mAdd(Add)
            'If LOG.Ret <> "OK" Then
            '    LOG.AddStepNameInf(DBConnName)
            '    Throw New Exception(LOG.Ret)
            'End If
            Me.ClearErr()
            Return Nothing
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function


    Public Function Remove(DBConnName As String) As String
        Dim LOG As New PigStepLog("Remove.DBConnName")
        Try
            LOG.StepName = "For Each"
            For Each oDBConnDef As DBConnDef In moList
                If oDBConnDef.DBConnName = DBConnName Then
                    LOG.AddStepNameInf(DBConnName)
                    moList.Remove(oDBConnDef)
                    Exit For
                End If
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function Remove(Index As Integer) As String
        Dim LOG As New PigStepLog("Remove.Index")
        Try
            LOG.StepName = "Index=" & Index.ToString
            moList.RemoveAt(Index)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    'Public Function AddOrGet(DBConnName As String, SQLServer As String, CurrDatabase As String, DBUser As String, DBUserPwd As String, Optional DBConnDesc As String = "") As DBConnDef
    '    Dim LOG As New PigStepLog("AddOrGet")
    '    Try
    '        If Me.IsItemExists(DBConnName) = True Then
    '            AddOrGet = Me.Item(DBConnName)
    '        Else
    '            AddOrGet = Me.Add(DBConnName, SQLServer, CurrDatabase, DBUser, DBUserPwd, DBConnDesc)
    '        End If
    '        Me.ClearErr()
    '    Catch ex As Exception
    '        Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '        Return Nothing
    '    End Try
    'End Function

    'Public Function AddOrGet(DBConnName As String, SQLServer As String, CurrDatabase As String, Optional DBConnDesc As String = "") As DBConnDef
    '    Dim LOG As New PigStepLog("AddOrGet")
    '    Try
    '        If Me.IsItemExists(DBConnName) = True Then
    '            AddOrGet = Me.Item(DBConnName)
    '        Else
    '            AddOrGet = Me.Add(DBConnName, SQLServer, CurrDatabase, DBConnDesc)
    '        End If
    '        Me.ClearErr()
    '    Catch ex As Exception
    '        Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '        Return Nothing
    '    End Try
    'End Function

    Public Function Clear() As String
        Try
            moList.Clear()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Clear", ex)
        End Try
    End Function

    Public ReadOnly Property IsChange As Boolean
        Get
            Try
                Return Me.Parent.fPigConfigApp.IsChange
            Catch ex As Exception
                Me.SetSubErrInf("IsChange", ex)
                Return False
            End Try
        End Get
    End Property

End Class

