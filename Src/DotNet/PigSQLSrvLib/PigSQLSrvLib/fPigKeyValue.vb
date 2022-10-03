'**********************************
'* Name: fPigKeyValue
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 豚豚键值
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 3/8/2022
'* 1.1		5/8/2022	Modify SortedList,mSaveKeyValue, add mSaveBodyToFile
'************************************
Imports PigToolsLiteLib
Friend Class fPigKeyValue
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.1.10"
    Public Enum HitCacheEnum
        Null = 0
        List = 1
        ShareMem = 2
        File = 3
    End Enum

    Private ReadOnly Property mPigFunc As New PigFunc
    Public ReadOnly Property CacheWorkDir As String
    Private ReadOnly Property mslValue As New SortedList
    Private ReadOnly Property mslCreateTime As New SortedList
    Private ReadOnly Property mslPigMD5 As New SortedList
    Public ReadOnly Property MaxWorkList As Integer

    Public Sub New(CacheWorkDir As String, Optional MaxWorkList As Integer = 100)
        MyBase.New(CLS_VERSION)
        Dim strRet As String = ""
        Try
            Me.CacheWorkDir = CacheWorkDir
            If MaxWorkList < 0 Then MaxWorkList = 100
            Me.MaxWorkList = MaxWorkList
            If Me.mPigFunc.IsFolderExists(Me.CacheWorkDir) = False Then
                strRet = Me.mPigFunc.CreateFolder(Me.CacheWorkDir)
                If strRet <> "OK" Then Throw New Exception("Failed to create directory " & Me.CacheWorkDir)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    Private Function mSaveKeyValueToList(KeyName As String, ByRef ValueBytes As Byte()) As String
        Dim LOG As New PigStepLog("mGetKeyValueFromList")
        Try
            LOG.StepName = "Check ValueBytes"
            If ValueBytes Is Nothing Then Throw New Exception("ValueBytes Is Nothing")
            If ValueBytes.Length = 0 Then Throw New Exception("ValueBytes Is empty")
            Dim intIndex As Integer
            '-------------
            LOG.StepName = "Get Value"
            intIndex = Me.mslValue.IndexOfKey(KeyName)
            If intIndex >= 0 Then
                LOG.StepName = "Remove Value"
                Me.mslValue.Remove(KeyName)
            End If
            LOG.StepName = "Add Value"
            Me.mslValue.Add(KeyName, ValueBytes)
            '-------------
            LOG.StepName = "Get CreateTime"
            intIndex = Me.mslCreateTime.IndexOfKey(KeyName)
            If intIndex >= 0 Then
                LOG.StepName = "Remove CreateTime"
                Me.mslCreateTime.Remove(KeyName)
            End If
            LOG.StepName = "Add CreateTime"
            Dim dteNow As Date = Now
            Me.mslCreateTime.Add(KeyName, dteNow)
            '-------------
            LOG.StepName = "Get PigMD5"
            intIndex = Me.mslPigMD5.IndexOfKey(KeyName)
            If intIndex >= 0 Then
                LOG.StepName = "Remove PigMD5"
                Me.mslPigMD5.Remove(KeyName)
            End If
            LOG.StepName = "Add PigMD5"
            Dim oPigMD5 As New PigMD5(ValueBytes)
            Me.mslPigMD5.Add(KeyName, oPigMD5.PigMD5)
            '-------------
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(KeyName)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetKeyValue(KeyName As String, ByRef TextValue As String, Optional TextType As PigText.enmTextType = PigText.enmTextType.UTF8, Optional CacheTimeSec As Integer = 60, Optional ByRef HitCache As HitCacheEnum = HitCacheEnum.Null) As String
        Dim LOG As New PigStepLog("mGetKeyValue")
        Try
            Dim abValue(0) As Byte
            LOG.StepName = "mGetKeyValue"
            LOG.Ret = Me.mGetKeyValue(KeyName, abValue, CacheTimeSec, HitCache)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "New PigText"
            Dim gtValue As New PigText(abValue, TextType)
            TextValue = gtValue.Text
            Return "OK"
        Catch ex As Exception
            TextValue = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mGetKeyValue(KeyName As String, ByRef ValueBytes As Byte(), Optional CacheTimeSec As Integer = 60, Optional ByRef HitCache As HitCacheEnum = HitCacheEnum.Null) As String
        Dim LOG As New PigStepLog("mGetKeyValue")
        Try
            Dim dteCreateTime As Date, bolIsNeedGetFromFile As Boolean = False, bolIsNeedGetFromShareMem As Boolean = False
            LOG.StepName = "mGetKeyValueFromList"
            LOG.Ret = Me.mGetKeyValueFromList(KeyName, ValueBytes, CacheTimeSec, dteCreateTime)
            If LOG.Ret <> "OK" Then
                bolIsNeedGetFromShareMem = True
            ElseIf dteCreateTime.AddSeconds(CacheTimeSec) < Now Then
                bolIsNeedGetFromShareMem = True
            Else
                HitCache = HitCacheEnum.List
            End If
            Dim strKeyName As String
            Dim abHead(0) As Byte, pbHead As PigBytes = Nothing, ValuePigMD5(0) As Byte, lngValueLen As Integer = 0, strValuePigMD5 As String = ""
            If bolIsNeedGetFromShareMem = True Then
                strKeyName = Me.mGetKeyNamePigMD5(KeyName)
                LOG.StepName = "GetShareMem(Head)"
                LOG.Ret = Me.mPigFunc.GetShareMem(strKeyName, abHead, dteCreateTime)
                If LOG.Ret <> "OK" Then
                    bolIsNeedGetFromFile = True
                ElseIf dteCreateTime.AddSeconds(CacheTimeSec) < Now Then
                    bolIsNeedGetFromFile = True
                Else
                    LOG.StepName = "New PigBytes(Head)"
                    pbHead = New PigBytes(abHead)
                    If pbHead.LastErr <> "" Then
                        bolIsNeedGetFromFile = True
                    Else
                        lngValueLen = pbHead.GetInt32Value
                        ValuePigMD5 = pbHead.GetBytesValue(16)
                        If lngValueLen <= 0 Then
                            bolIsNeedGetFromFile = True
                        Else
                            strValuePigMD5 = Me.mPigFunc.GetPigMD5OrMD5(ValuePigMD5)
                            If Len(strValuePigMD5) <> 32 Then
                                bolIsNeedGetFromFile = True
                            Else
                                LOG.StepName = "GetShareMem(Value)"
                                LOG.Ret = Me.mPigFunc.GetShareMem(strValuePigMD5, ValueBytes, dteCreateTime)
                                If LOG.Ret <> "OK" Then
                                    bolIsNeedGetFromFile = True
                                ElseIf ValueBytes Is Nothing Then
                                    bolIsNeedGetFromFile = True
                                ElseIf ValueBytes.Length <> lngValueLen Then
                                    bolIsNeedGetFromFile = True
                                Else
                                    Dim oPigMD5 As New PigMD5(ValueBytes)
                                    If oPigMD5.PigMD5 <> strValuePigMD5 Then
                                        bolIsNeedGetFromFile = True
                                    Else
                                        LOG.StepName = "mSaveKeyValueToList"
                                        LOG.Ret = Me.mSaveKeyValueToList(KeyName, ValueBytes)
                                        If LOG.Ret <> "OK" Then Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
                                        HitCache = HitCacheEnum.ShareMem
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If bolIsNeedGetFromFile = True Then
                strKeyName = Me.mGetKeyNamePigMD5(KeyName)
                LOG.StepName = "mGetHeadFromFile(Head)"
                LOG.Ret = Me.mGetHeadFromFile(strKeyName, abHead, dteCreateTime)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                If dteCreateTime.AddSeconds(CacheTimeSec) < Now Then Throw New Exception("Data expiration")
                LOG.StepName = "New PigBytes(Head)"
                pbHead = New PigBytes(abHead)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                LOG.StepName = "GetInt32Value"
                lngValueLen = pbHead.GetInt32Value
                If lngValueLen <= 0 Then Throw New Exception("No data")
                LOG.StepName = "GetPigMD5OrMD5"
                ValuePigMD5 = pbHead.GetBytesValue(16)
                strValuePigMD5 = Me.mPigFunc.GetPigMD5OrMD5(ValuePigMD5)
                If Len(strValuePigMD5) <> 32 Then Throw New Exception("Invalid data")
                LOG.StepName = "mGetValueFromFile(Value)"
                LOG.Ret = Me.mGetValueFromFile(strValuePigMD5, ValueBytes)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strValuePigMD5)
                    Throw New Exception(LOG.Ret)
                End If
                If ValueBytes Is Nothing Then Throw New Exception("ValueBytes Is Nothing")
                If ValueBytes.Length <> lngValueLen Then Throw New Exception("Length mismatch")
                Dim oPigMD5 As New PigMD5(ValueBytes)
                If oPigMD5.PigMD5 <> strValuePigMD5 Then Throw New Exception("PigMD5 mismatch")
                oPigMD5 = Nothing
                Dim bolIsSaveToList As Boolean = False
                If Me.IsWindows = True Then
                    LOG.StepName = "SaveShareMem(Head)"
                    LOG.Ret = Me.mPigFunc.SaveShareMem(strKeyName, abHead)
                    If LOG.Ret <> "OK" Then
                        bolIsSaveToList = True
                    Else
                        LOG.StepName = "SaveShareMem(Body)"
                        Me.mPigFunc.SaveShareMem(strValuePigMD5, ValueBytes)
                        If LOG.Ret <> "OK" Then bolIsSaveToList = True
                    End If
                Else
                    bolIsSaveToList = True
                End If
                If bolIsSaveToList = True Then
                    LOG.StepName = "mSaveKeyValueToList"
                    LOG.Ret = Me.mSaveKeyValueToList(KeyName, ValueBytes)
                    If LOG.Ret <> "OK" Then Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
                End If
                HitCache = HitCacheEnum.File
            End If
            LOG.StepName = "New PigBytes(ValueBytes)"
            Dim oPigBytes As New PigBytes(ValueBytes)
            LOG.StepName = "UnCompress"
            LOG.Ret = oPigBytes.UnCompress()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Main to ValueBytes"
            ReDim ValueBytes(0)
            ValueBytes = oPigBytes.Main
            oPigBytes = Nothing
            Return "OK"
        Catch ex As Exception
            ReDim ValueBytes(0)
            LOG.AddStepNameInf(KeyName)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Private Function mGetKeyValueFromList(KeyName As String, ByRef ValueBytes As Byte(), CacheTimeSec As Integer, ByRef CreateTime As Date) As String
        Dim LOG As New PigStepLog("mGetKeyValueFromList")
        Try
            LOG.StepName = "Get CreateTime"
            If Me.mslCreateTime.IndexOfKey(KeyName) < 0 Then Throw New Exception("No data")
            CreateTime = Me.mslCreateTime.Item(KeyName)
            If CreateTime.AddSeconds(CacheTimeSec) < Now Then Throw New Exception("Data expiration")
            LOG.StepName = "Get PigMD5"
            If Me.mslPigMD5.IndexOfKey(KeyName) < 0 Then Throw New Exception("No data")
            Dim strPigMD5 As String = Me.mslPigMD5.Item(KeyName)
            If Len(strPigMD5) <> 32 Then Throw New Exception("Invalid data")
            LOG.StepName = "Get Value"
            If Me.mslValue.IndexOfKey(KeyName) < 0 Then Throw New Exception("No data")
            ValueBytes = Me.mslValue.Item(KeyName)
            LOG.StepName = "Check PigMD5"
            Dim oPigMD5 As New PigMD5(ValueBytes)
            If oPigMD5.PigMD5 <> strPigMD5 Then Throw New Exception("Data mismatch")
            Return "OK"
        Catch ex As Exception
            ReDim ValueBytes(0)
            CreateTime = Date.MinValue
            LOG.AddStepNameInf(KeyName)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function SaveKeyValue(KeyName As String, SaveText As String, Optional TextType As PigText.enmTextType = PigText.enmTextType.UTF8) As String
        Try
            Dim ptMain As New PigText(SaveText, TextType)
            Dim strRet As String = Me.mSaveKeyValue(KeyName, ptMain.TextBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            ptMain = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SaveKeyValue", ex)
        End Try
    End Function

    Private Function mGetKeyNamePigMD5(KeyName As String) As String
        Try
            mGetKeyNamePigMD5 = ""
            Dim strRet As String = Me.mPigFunc.GetTextPigMD5("~PigShareMem.(" & KeyName & "#>PigShareMem,>", PigMD5.enmTextType.UTF8, mGetKeyNamePigMD5)
            If strRet <> "OK" Then Throw New Exception(strRet)
        Catch ex As Exception
            Me.SetSubErrInf("mGetKeyNamePigMD5", ex)
            Return ""
        End Try
    End Function



    Private Function mSaveKeyValue(KeyName As String, DataBytes As Byte()) As String
        Dim LOG As New PigStepLog("mSaveKeyValue")
        Try
            LOG.StepName = "Check DataBytes"
            If DataBytes Is Nothing Then Throw New Exception("DataBytes Is Nothing")
            If DataBytes.Length = 0 Then Throw New Exception("DataBytes Is empty")
            LOG.StepName = "mGetKeyNamePigMD5"
            Dim strKeyName As String = Me.mGetKeyNamePigMD5(KeyName)
            If strKeyName = "" Then Throw New Exception("Unable to get")
            '---------
            LOG.StepName = "New PigBytes"
            Dim pbMain As New PigBytes(DataBytes)
            If pbMain.LastErr <> "" Then Throw New Exception(pbMain.LastErr)
            LOG.StepName = "Compress"
            LOG.Ret = pbMain.Compress()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "mSaveBodyToFile"
            LOG.Ret = Me.mSaveBodyToFile(pbMain.PigMD5, pbMain.Main)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            '---------
            LOG.StepName = "mSaveHeadToFile"
            LOG.Ret = Me.mSaveHeadToFile(strKeyName, pbMain)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            '---------
            pbMain = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mGetHeadFromFile(KeyNamePigMD5 As String, ByRef HeadBytes As Byte(), ByRef CreateTime As Date) As String
        Dim LOG As New PigStepLog("mGetValueFromFile")
        Dim strFilePath As String = Me.CacheWorkDir & Me.OsPathSep & KeyNamePigMD5
        Try
            LOG.StepName = "New PigFile"
            Dim oPigFile As New PigFile(strFilePath)
            LOG.StepName = "LoadFile"
            LOG.Ret = oPigFile.LoadFile
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "GbMain.Main"
            HeadBytes = oPigFile.GbMain.Main
            CreateTime = oPigFile.UpdateTime
            oPigFile = Nothing
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strFilePath)
            ReDim HeadBytes(0)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mGetValueFromFile(ValuePigMD5 As String, ByRef OutData As Byte()) As String
        Dim LOG As New PigStepLog("mGetValueFromFile")
        Dim strFilePath As String = Me.CacheWorkDir & Me.OsPathSep & ValuePigMD5
        Try
            LOG.StepName = "New PigFile"
            Dim oPigFile As New PigFile(strFilePath)
            LOG.StepName = "LoadFile"
            LOG.Ret = oPigFile.LoadFile
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "GbMain.Main"
            OutData = oPigFile.GbMain.Main
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strFilePath)
            ReDim OutData(0)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Private Function mSaveHeadToFile(KeyNamePigMD5 As String, ByRef PbBody As PigBytes) As String
        Dim LOG As New PigStepLog("mSaveHeadToFile")
        Dim strFilePath As String = Me.CacheWorkDir & Me.OsPathSep & KeyNamePigMD5
        Try
            LOG.StepName = "New PigBytes"
            Dim pbHead As New PigBytes
            With pbHead
                LOG.StepName = "pbHead.SetValue"
                .SetValue(PbBody.Main.Length)
                .SetValue(PbBody.PigMD5Bytes)
                If .LastErr <> "" Then Throw New Exception(.LastErr)
            End With
            LOG.StepName = "New PigFile"
            Dim oPigFile As New PigFile(strFilePath)
            oPigFile.GbMain = pbHead
            LOG.StepName = "SaveFile"
            LOG.Ret = oPigFile.SaveFile()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            oPigFile = Nothing
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strFilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mSaveBodyToFile(ValuePigMD5 As String, ByRef SaveData As Byte()) As String
        Dim LOG As New PigStepLog("mSaveBodyToFile")
        Try
            Dim bolIsSave As Boolean = True
            Dim strFilePath As String = Me.CacheWorkDir & Me.OsPathSep & ValuePigMD5
            If Me.mPigFunc.IsFileExists(strFilePath) = True Then
                Dim strPigMD5 As String = ""
                LOG.StepName = "GetFilePigMD5"
                LOG.Ret = Me.mPigFunc.GetFilePigMD5(strFilePath, strPigMD5)
                If LOG.Ret = "OK" Then
                    If ValuePigMD5 = strPigMD5 Then
                        bolIsSave = False
                    End If
                End If
            End If
            If bolIsSave = True Then
                LOG.StepName = "New PigFile"
                Dim oPigFile As New PigFile(strFilePath)
                oPigFile.GbMain = New PigBytes(SaveData)
                LOG.StepName = "SaveFile"
                LOG.Ret = oPigFile.SaveFile()
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strFilePath)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
