'**********************************
'* Name: SQLSrvKeyValue
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigKeyValue of SQL Server
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.6
'* Create Time: 1/10/2022
'* 1.1  1/10/2022   Modify New,mAddTableCol, add mNew,RefDBConn,SaveKeyValue
'* 1.2  2/10/2022   Add mGetKeyValue,GetKeyValue,mCreateTableKeyValueHeadInf,mCreateTableKeyValueBodyInf,mSaveBodyToDB,mSaveHeadToDB
'* 1.3  3/10/2022   Modify mGetKeyValue
'* 1.5  4/10/2022   Add RemoveKeyValue,mRemoveKeyValue, modify mGetBodyFromDB,mGetHeadFromDB,mGetKeyValue
'* 1.6  21/7/2024  Modify PigFunc to PigFuncLite
'**********************************
Imports System.Data
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
#Else
Imports Microsoft.Data.SqlClient
#End If
Imports PigToolsLiteLib
''' <summary>
''' 用于SQL Server 的键值|PigKeyValue of SQL Server
''' </summary>
Public Class SQLSrvKeyValue
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1." & "6" & ".10"

    Private Property mConnSQLSrv As ConnSQLSrv
    Private Property mPigFunc As New PigFuncLite
    Private Property mPigKeyValue As PigKeyValue
    Private Property mSeowEnc As SeowEnc

    Public Sub New(ConnSQLSrv As ConnSQLSrv, CacheWorkDir As String, Optional MaxWorkList As Integer = 100)
        MyBase.New(CLS_VERSION)
        Me.mNew(ConnSQLSrv, CacheWorkDir, True, MaxWorkList)
    End Sub

    Public Sub New(ConnSQLSrv As ConnSQLSrv, CacheWorkDir As String, IsCompress As Boolean, Optional MaxWorkList As Integer = 100)
        MyBase.New(CLS_VERSION)
        Me.mNew(ConnSQLSrv, CacheWorkDir, IsCompress, MaxWorkList)
    End Sub

    Private mIsDBReady As Boolean
    Public ReadOnly Property IsDBReady As Boolean
        Get
            Return Me.mIsDBReady
        End Get
    End Property
    ''' <summary>
    ''' Refresh database connection|刷新数据库连接
    ''' </summary>
    ''' <returns></returns>
    Public Function RefDBConn() As String
        Dim LOG As New PigStepLog("RefDBConn")
        Try
            If Me.mConnSQLSrv.IsDBConnReady = False Then
                LOG.StepName = "OpenOrKeepActive"
                Me.mConnSQLSrv.ClearErr()
                Me.mConnSQLSrv.OpenOrKeepActive()
                If Me.mConnSQLSrv.LastErr <> "" Then Throw New Exception(Me.mConnSQLSrv.LastErr)
            End If
            Dim oSQLSrvTools As New SQLSrvTools(Me.mConnSQLSrv)
            LOG.StepName = "IsDBObjExists"
            If oSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptKeyValueHeadInf") = False Then
                LOG.StepName = "mCreateTableKeyValueHeadInf"
                LOG.Ret = mCreateTableKeyValueHeadInf()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End If
            If oSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptKeyValueBodyInf") = False Then
                LOG.StepName = "mCreateTableKeyValueBodyInf"
                LOG.Ret = mCreateTableKeyValueBodyInf()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End If
            oSQLSrvTools = Nothing
            Me.mIsDBReady = True
            Return "OK"
        Catch ex As Exception
            Me.mIsDBReady = False
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Sub mNew(ConnSQLSrv As ConnSQLSrv, CacheWorkDir As String, IsCompress As Boolean, Optional MaxWorkList As Integer = 100)
        Dim LOG As New PigStepLog("mNew")
        Try
            LOG.StepName = "New PigKeyValue"
            Me.mPigKeyValue = New PigKeyValue(CacheWorkDir, IsCompress, MaxWorkList)
            If Me.mPigKeyValue.LastErr <> "" Then Throw New Exception(Me.mPigKeyValue.LastErr)
            LOG.StepName = "Set ConnSQLSrv"
            Me.mConnSQLSrv = ConnSQLSrv
            LOG.StepName = "RefDBConn"
            LOG.Ret = Me.RefDBConn()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            If Me.mPigKeyValue.IsCompress = True Then
                Me.mSeowEnc = New SeowEnc(SeowEnc.EmnComprssType.AutoComprss)
                Dim oPigMD5 As New PigMD5(Me.mConnSQLSrv.CurrDatabase, PigMD5.enmTextType.UTF8)
                Dim abEncKey(23) As Byte
                For i = 0 To 15
                    abEncKey(i) = oPigMD5.PigMD5Bytes(i)
                Next
                For i = 16 To 23
                    abEncKey(i) = oPigMD5.MD5Bytes(i - 16)
                Next
                oPigMD5 = Nothing
                LOG.StepName = "SeowEnc.LoadEncKey"
                LOG.Ret = Me.mSeowEnc.LoadEncKey(abEncKey)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                Me.mSeowEnc.IsRandAdd = False   '这样生成的密文会固定，这样才能作为缓存
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub
    Private Function mCreateTableKeyValueHeadInf() As String
        Dim LOG As New PigStepLog("mCreateTableKeyValueHeadInf")
        Dim strSQL As String = ""
        Try
            Dim strTabName As String = ""
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "CREATE TABLE dbo._ptKeyValueHeadInf(")
                .AddMultiLineText(strSQL, "KeyName varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",BodyLen int NOT NULL", 1)
                .AddMultiLineText(strSQL, ",BodyPigMD5 varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
                .AddMultiLineText(strSQL, "CONSTRAINT PK_KeyValueHeadInf PRIMARY KEY CLUSTERED(KeyName)", 1)
                .AddMultiLineText(strSQL, ")")
                .AddMultiLineText(strSQL, "CREATE INDEX Idx_ptKeyValueHeadInfCreateTime ON dbo._ptKeyValueHeadInf(CreateTime)")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                LOG.StepName = "ExecuteNonQuery"
                LOG.Ret = .ExecuteNonQuery()
                If LOG.Ret <> "OK" Then
                    Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
                    Throw New Exception(LOG.Ret)
                End If
            End With
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mCreateTableKeyValueBodyInf() As String
        Dim LOG As New PigStepLog("mCreateTableKeyValueBodyInf")
        Dim strSQL As String = ""
        Try
            Dim strTabName As String = ""
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "CREATE TABLE dbo._ptKeyValueBodyInf(")
                .AddMultiLineText(strSQL, "BodyPigMD5 varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",BodyData varchar(max) NOT NULL DEFAULT ('')", 1)
                .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
                .AddMultiLineText(strSQL, "CONSTRAINT PK_KeyValueBodyInf PRIMARY KEY CLUSTERED(BodyPigMD5)", 1)
                .AddMultiLineText(strSQL, ")")
                .AddMultiLineText(strSQL, "CREATE INDEX IdxKeyValueBodyInfCreateTime ON dbo._ptKeyValueBodyInf(CreateTime)")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                LOG.StepName = "ExecuteNonQuery"
                LOG.Ret = .ExecuteNonQuery()
                If LOG.Ret <> "OK" Then
                    Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
                    Throw New Exception(LOG.Ret)
                End If
            End With
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    'Private Function mAddTableCol() As String
    '    Const SUB_NAME As String = "mAddTableCol"
    '    Dim strStepName As String = "", strRet As String = ""
    '    Try
    '        Dim strTabName As String = ""
    '        Dim strSQL As String = ""
    '        With Me.mPigFunc
    '            .AddMultiLineText(strSQL, "IF NOT EXISTS(SELECT 1 FROM syscolumns c JOIN sysobjects o ON c.id=o.id AND o.xtype='U' AND o.uid=1 WHERE o.name='_ptKeyValueHeadInf' AND c.name='HeadData')")
    '            .AddMultiLineText(strSQL, "BEGIN")
    '            .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptKeyValueInf ADD HeadData varchar(256) NOT NULL DEFAULT ('')", 1)
    '            .AddMultiLineText(strSQL, "END")
    '            .AddMultiLineText(strSQL, "IF NOT EXISTS(SELECT 1 FROM syscolumns c JOIN sysobjects o ON c.id=o.id AND o.xtype='U' AND o.uid=1 WHERE o.name='_ptKeyValueBodyInf' AND c.name='BodyData')")
    '            .AddMultiLineText(strSQL, "BEGIN")
    '            .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptKeyValueInf ADD BodyData varchar(max) NOT NULL DEFAULT ('')", 1)
    '            .AddMultiLineText(strSQL, "END")
    '        End With
    '        strStepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            strStepName = "ExecuteNonQuery"
    '            strRet = .ExecuteNonQuery()
    '            If strRet <> "OK" Then
    '                Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
    '                Throw New Exception(strRet)
    '            End If
    '        End With
    '        Return "OK"
    '    Catch ex As Exception
    '        Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
    '    End Try
    'End Function

    ''' <summary>
    ''' Save Key Value|保存键值
    ''' </summary>
    ''' <param name="KeyName">Key Value Name|键值名称</param>
    ''' <param name="DataBytes">Byte array for saving data|保存数据的字节数组</param>
    ''' <returns>Execution result: OK indicates success, and others are error messages|执行结果，OK表示成功，其他为错误信息</returns>
    Public Function SaveKeyValue(KeyName As String, DataBytes As Byte()) As String
        Return Me.mSaveKeyValue(KeyName, DataBytes)
    End Function

    ''' <summary>
    ''' Save Key Value|保存键值
    ''' </summary>
    ''' <param name="KeyName">Key Value Name|键值名称</param>
    ''' <param name="Base64SaveText">Base64 of byte array for saving data|保存数据的字节数组的Base64</param>
    ''' <returns>Execution result: OK indicates success, and others are error messages|执行结果，OK表示成功，其他为错误信息</returns>
    Public Function SaveKeyValue(KeyName As String, Base64SaveText As String) As String
        Try
            Dim pbMain As New PigBytes(Base64SaveText)
            If pbMain.LastErr <> "" Then Throw New Exception(pbMain.LastErr)
            Dim strRet As String = Me.mSaveKeyValue(KeyName, pbMain.Main)
            If strRet <> "OK" Then Throw New Exception(strRet)
            pbMain = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SaveKeyValue", ex)
        End Try
    End Function

    ''' <summary>
    ''' Save Key Value|保存键值
    ''' </summary>
    ''' <param name="KeyName">Key Value Name|键值名称</param>
    ''' <param name="SaveText">Text to save data|保存数据的文本</param>
    ''' <param name="TextType">Text Type|文本类型</param>
    ''' <returns>Execution result: OK indicates success, and others are error messages|执行结果，OK表示成功，其他为错误信息</returns>
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
            Select Case Len(KeyName)
                Case = 0
                    Throw New Exception("KeyName not specified")
                Case 1 To 128
                Case > 128
                    Throw New Exception("KeyName length cannot exceed 128")
            End Select
            '---------
            If Me.mPigKeyValue.IsCompress = True Then
                Dim abData(0) As Byte
                LOG.StepName = "SeowEnc.Encrypt"
                LOG.Ret = Me.mSeowEnc.Encrypt(DataBytes, abData)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                DataBytes = abData
            End If
            LOG.StepName = "New PigBytes"
            Dim pbMain As New PigBytes(DataBytes)
            If pbMain.LastErr <> "" Then Throw New Exception(pbMain.LastErr)
            LOG.StepName = "mSaveBodyToDB"
            LOG.Ret = Me.mSaveBodyToDB(pbMain.PigMD5, pbMain.Main)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            '---------
            LOG.StepName = "mSaveHeadToDB"
            LOG.Ret = Me.mSaveHeadToDB(KeyName, pbMain)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            '---------
            Me.mRemoveKeyValue(KeyName, False)
            '---------
            pbMain = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    ''' <summary>
    ''' Get Key Value|获取键值
    ''' </summary>
    ''' <param name="KeyName">Key Value Name|键值名称</param>
    ''' <param name="ValueBytes">Byte array obtained|获取到的字节数组</param>
    ''' <param name="CacheTimeSec">Seconds of cache time|缓存时间的秒数</param>
    ''' <param name="HitCache">Hit Cache Level|命中缓存级别</param>
    ''' <returns>Execution result: OK indicates success, and others are error messages|执行结果，OK表示成功，其他为错误信息</returns>
    Public Function GetKeyValue(KeyName As String, ByRef ValueBytes As Byte(), Optional CacheTimeSec As Integer = 60, Optional ByRef HitCache As PigKeyValue.HitCacheEnum = PigKeyValue.HitCacheEnum.Null) As String
        Return Me.mGetKeyValue(KeyName, ValueBytes, CacheTimeSec, HitCache)
    End Function

    ''' <summary>
    ''' Get Key Value|获取键值
    ''' </summary>
    ''' <param name="KeyName">Key Value Name|键值名称</param>
    ''' <param name="Base64Value">Base64 of the obtained byte array|获取到的字节数组的Base64</param>
    ''' <param name="CacheTimeSec">Seconds of cache time|缓存时间的秒数</param>
    ''' <param name="HitCache">Hit Cache Level|命中缓存级别</param>
    ''' <returns>Execution result: OK indicates success, and others are error messages|执行结果，OK表示成功，其他为错误信息</returns>
    Public Function GetKeyValue(KeyName As String, ByRef Base64Value As String, Optional CacheTimeSec As Integer = 60, Optional ByRef HitCache As PigKeyValue.HitCacheEnum = PigKeyValue.HitCacheEnum.Null) As String
        Dim LOG As New PigStepLog("GetKeyValue")
        Try
            Dim abValue(0) As Byte
            LOG.StepName = "mGetKeyValue"
            LOG.Ret = Me.mGetKeyValue(KeyName, abValue, CacheTimeSec, HitCache)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "ToBase64String"
            Base64Value = Convert.ToBase64String(abValue)
            Return "OK"
        Catch ex As Exception
            Base64Value = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Get Key Value|获取键值
    ''' </summary>
    ''' <param name="KeyName">Key Value Name|键值名称</param>
    ''' <param name="TextValue">Text obtained|获取到的文本</param>
    ''' <param name="TextType">Text Type|文本类型</param>
    ''' <param name="CacheTimeSec">Seconds of cache time|缓存时间的秒数</param>
    ''' <param name="HitCache">Hit Cache Level|命中缓存级别</param>
    ''' <returns>Execution result: OK indicates success, and others are error messages|执行结果，OK表示成功，其他为错误信息</returns>
    Public Function GetKeyValue(KeyName As String, ByRef TextValue As String, Optional TextType As PigText.enmTextType = PigText.enmTextType.UTF8, Optional CacheTimeSec As Integer = 60, Optional ByRef HitCache As PigKeyValue.HitCacheEnum = PigKeyValue.HitCacheEnum.Null) As String
        Dim LOG As New PigStepLog("GetKeyValue")
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

    Private Function mGetKeyValue(KeyName As String, ByRef ValueBytes As Byte(), Optional CacheTimeSec As Integer = 60, Optional ByRef HitCache As PigKeyValue.HitCacheEnum = PigKeyValue.HitCacheEnum.Null) As String
        Dim LOG As New PigStepLog("mGetKeyValue")
        Try
            Dim dteCreateTime As Date, bolIsNeedGetFromDB As Boolean = False
            LOG.Ret = Me.mPigKeyValue.GetKeyValue(KeyName, ValueBytes, CacheTimeSec, HitCache)
            If LOG.Ret <> "OK" Then bolIsNeedGetFromDB = True
            If bolIsNeedGetFromDB = True Then
                Dim lngBodyLen As Long, strBodyPigMD5 As String = ""
                LOG.StepName = "mGetHeadFromDB"
                LOG.Ret = Me.mGetHeadFromDB(KeyName, lngBodyLen, strBodyPigMD5, dteCreateTime)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                If dteCreateTime.AddSeconds(CacheTimeSec) < Now Then Throw New Exception("Data expiration")
                If Len(strBodyPigMD5) <> 32 Then Throw New Exception("Invalid data")
                LOG.StepName = "mGetBodyFromDB"
                LOG.Ret = Me.mGetBodyFromDB(strBodyPigMD5, lngBodyLen, ValueBytes)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strBodyPigMD5)
                    Throw New Exception(LOG.Ret)
                End If
                If ValueBytes Is Nothing Then Throw New Exception("ValueBytes Is Nothing")
                Dim oPigMD5 As New PigMD5(ValueBytes)
                If oPigMD5.PigMD5 <> strBodyPigMD5 Then Throw New Exception("PigMD5 mismatch")
                oPigMD5 = Nothing
                Dim abValue(0) As Byte
                If Me.mPigKeyValue.IsCompress = True Then
                    LOG.StepName = "SeowEnc.Decrypt"
                    LOG.Ret = Me.mSeowEnc.Decrypt(ValueBytes, abValue)
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                    ReDim ValueBytes(abValue.Length - 1)
                    abValue.CopyTo(ValueBytes, 0)
                Else
                    ReDim abValue(ValueBytes.Length - 1)
                    ValueBytes.CopyTo(abValue, 0)
                End If
                HitCache = PigKeyValue.HitCacheEnum.DB
                LOG.StepName = "mPigKeyValue.SaveKeyValue"
                LOG.Ret = Me.mPigKeyValue.SaveKeyValue(KeyName, abValue)
            End If
            Return "OK"
        Catch ex As Exception
            ReDim ValueBytes(0)
            LOG.AddStepNameInf(KeyName)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mSaveBodyToDB(ValuePigMD5 As String, ByRef SaveData As Byte()) As String
        Dim LOG As New PigStepLog("mSaveBodyToDB")
        Dim strSQL As String = ""
        Try
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "IF NOT EXISTS(SELECT TOP 1 1 FROM dbo._ptKeyValueBodyInf WHERE BodyPigMD5=@BodyPigMD5)")
                .AddMultiLineText(strSQL, "INSERT INTO dbo._ptKeyValueBodyInf(BodyPigMD5,BodyData)VALUES(@BodyPigMD5,@BodyData)", 1)
                .AddMultiLineText(strSQL, "ELSE")
                .AddMultiLineText(strSQL, "UPDATE dbo._ptKeyValueBodyInf SET BodyData=@BodyData", 1)
                .AddMultiLineText(strSQL, "WHERE BodyPigMD5=@BodyPigMD5 AND  BodyData!=@BodyData", 1)
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                .AddPara("@BodyPigMD5", SqlDbType.VarChar, 32)
                .AddPara("@BodyData", SqlDbType.VarChar, -1)
                .ParaValue("@BodyPigMD5") = ValuePigMD5
                .ParaValue("@BodyData") = Convert.ToBase64String(SaveData)
                LOG.StepName = "ExecuteNonQuery"
                If Me.IsDebug = True Then LOG.AddStepNameInf(.DebugStr)
                LOG.Ret = .ExecuteNonQuery
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End With
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mSaveHeadToDB(KeyName As String, ByRef PbBody As PigBytes) As String
        Dim LOG As New PigStepLog("mSaveBodyToDB")
        Dim strSQL As String = ""
        Try
            If PbBody Is Nothing Then Throw New Exception("PbBody Is Nothing")
            If PbBody.Main Is Nothing Then Throw New Exception("PbBody.Main Is Nothing")
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "IF NOT EXISTS(SELECT TOP 1 1 FROM dbo._ptKeyValueHeadInf WHERE KeyName=@KeyName)")
                .AddMultiLineText(strSQL, "INSERT INTO dbo._ptKeyValueHeadInf(KeyName,BodyLen,BodyPigMD5)VALUES(@KeyName,@BodyLen,@BodyPigMD5)", 1)
                .AddMultiLineText(strSQL, "ELSE")
                .AddMultiLineText(strSQL, "UPDATE dbo._ptKeyValueHeadInf SET BodyLen=@BodyLen,BodyPigMD5=@BodyPigMD5,CreateTime=GetDate()", 1)
                .AddMultiLineText(strSQL, "WHERE KeyName=@KeyName", 1)
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                .AddPara("@KeyName", SqlDbType.VarChar, 32)
                .AddPara("@BodyLen", SqlDbType.Int)
                .AddPara("@BodyPigMD5", SqlDbType.VarChar, 32)
                .ParaValue("@KeyName") = KeyName
                .ParaValue("@BodyLen") = PbBody.Main.Length
                .ParaValue("@BodyPigMD5") = PbBody.PigMD5
                LOG.StepName = "ExecuteNonQuery"
                If Me.IsDebug = True Then LOG.AddStepNameInf(.DebugStr)
                LOG.Ret = .ExecuteNonQuery
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End With
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mGetHeadFromDB(KeyNamePigMD5 As String, ByRef BodyLen As Long, ByRef BodyPigMD5 As String, ByRef CreateTime As Date) As String
        Dim LOG As New PigStepLog("mGetHeadFromDB")
        Dim strSQL As String = ""
        Try
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "SELECT TOP 1 BodyLen,BodyPigMD5,CreateTime FROM dbo._ptKeyValueHeadInf WITH (NOLOCK) WHERE KeyName=@KeyName")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                .AddPara("@KeyName", SqlDbType.VarChar, 32)
                .ParaValue("@KeyName") = KeyNamePigMD5
                LOG.StepName = "Execute"
                If Me.IsDebug = True Then LOG.AddStepNameInf(.DebugStr)
                Dim rsMain As Recordset = .Execute()
                If .LastErr <> "" Then
                    Throw New Exception(.LastErr)
                ElseIf rsMain Is Nothing Then
                    Throw New Exception("rsMain Is Nothing")
                ElseIf rsMain.EOF = True Then
                    Throw New Exception("Not data")
                Else
                    LOG.StepName = "Set value"
                    BodyLen = rsMain.Fields.Item("BodyLen").IntValue
                    BodyPigMD5 = rsMain.Fields.Item("BodyPigMD5").StrValue
                    CreateTime = rsMain.Fields.Item("CreateTime").DateValue
                End If
            End With
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strSQL)
            BodyLen = -1
            BodyPigMD5 = ""
            CreateTime = Date.MinValue
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mGetBodyFromDB(BodyPigMD5 As String, BodyLen As Long, ByRef BodyData As Byte()) As String
        Dim LOG As New PigStepLog("mGetBodyFromDB")
        Dim strSQL As String = ""
        Try
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "SELECT TOP 1 BodyData FROM dbo._ptKeyValueBodyInf WITH (NOLOCK) WHERE BodyPigMD5=@BodyPigMD5")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                .AddPara("@BodyPigMD5", SqlDbType.VarChar, 32)
                .ParaValue("@BodyPigMD5") = BodyPigMD5
                LOG.StepName = "Execute"
                If Me.IsDebug = True Then LOG.AddStepNameInf(.DebugStr)
                Dim rsMain As Recordset = .Execute()
                If .LastErr <> "" Then
                    Throw New Exception(.LastErr)
                ElseIf rsMain Is Nothing Then
                    Throw New Exception("rsMain Is Nothing")
                ElseIf rsMain.EOF = True Then
                    Throw New Exception("Not data")
                Else
                    LOG.StepName = "Set value"
                    Dim strBodyDataBase64 As String = rsMain.Fields.Item("BodyData").StrValue
                    Dim oPigBytes As New PigBytes(strBodyDataBase64)
                    If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
                    If BodyLen <> oPigBytes.Main.Length Then Throw New Exception("The data length of the Body does not match")
                    ReDim BodyData(BodyLen - 1)
                    oPigBytes.Main.CopyTo(BodyData, 0)
                    oPigBytes = Nothing
                End If
            End With
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strSQL)
            Return BodyData(0)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Delete Key Value|删除键值
    ''' </summary>
    ''' <param name="KeyName">Key Value Name|键值名称</param>
    ''' <returns>Execution result: OK indicates success, and others are error messages|执行结果，OK表示成功，其他为错误信息</returns>
    Public Function RemoveKeyValue(KeyName As String) As String
        Return Me.mRemoveKeyValue(KeyName， True)
    End Function

    Private Function mRemoveKeyValue(KeyName As String, IsIncDB As Boolean) As String
        Dim LOG As New PigStepLog("mRemoveKeyValue")
        Try
            Dim strError As String = ""
            Dim strKeyNamePigMD5 As String = Me.mGetKeyNamePigMD5(KeyName)
            If IsIncDB = True Then
                Dim strSQL As String = "UPDATE dbo._ptKeyValueHeadInf SET CreateTime='1/1/1900' WHERE KeyName=@KeyName"
                LOG.StepName = "New CmdSQLSrvText"
                Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
                With oCmdSQLSrvText
                    .ActiveConnection = Me.mConnSQLSrv.Connection
                    .AddPara("@KeyName", SqlDbType.VarChar, 32)
                    .ParaValue("@KeyName") = KeyName
                    LOG.StepName = "ExecuteNonQuery"
                    If Me.IsDebug = True Then LOG.AddStepNameInf(.DebugStr)
                    LOG.Ret = .ExecuteNonQuery
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                End With
            End If
            '---------
            LOG.StepName = "RemoveKeyValue"
            LOG.Ret = Me.mPigKeyValue.RemoveKeyValue(KeyName)
            If LOG.Ret <> "OK" Then strError &= LOG.StepLogInf
            '---------
            If strError <> "" Then Throw New Exception(strError)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
