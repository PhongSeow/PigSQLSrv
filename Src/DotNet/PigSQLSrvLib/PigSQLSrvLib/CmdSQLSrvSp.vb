﻿'**********************************
'* Name: CmdSQLSrvSp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: SqlCommand for SQL Server StoredProcedure
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.16
'* Create Time: 17/4/2021
'* 1.0.2	18/4/2021	Modify ActiveConnection
'* 1.0.3	24/4/2021	Add mAdoDataType
'* 1.0.4	25/4/2021	Modify New
'* 1.0.5	28/4/2021	Add ActiveConnection,AddPara,ParaValue,Execute
'* 1.0.6	16/5/2021	SQLSrvDataTypeEnum move to ConnSQLSrv, Modify Execute,ParaValue,ActiveConnection
'* 1.0.7	12/6/2021	Move to PigSQLSrvLib
'* 1.0.8	17/7/2021	Add DebugStr,mSQLStr,Modify New
'* 1.0.9	19/7/2021	Modify Execute
'* 1.0.10	28/7/2021	Modify DebugStr
'* 1.0.11	1/8/2021	Modify DebugStr
'* 1.1		29/8/2021   Add support for .net core
'* 1.2		4/9/2021	Add RecordsAffected,ExecuteNonQuery, Modify Execute
'* 1.3		24/9/2021	Add KeyName,CacheQuery
'* 1.4		6/10/2021	Modify CacheQuery
'* 1.5		8/10/2021	Modify CacheQuery
'* 1.6		15/12/2021	Modify CacheQuery, and Rewrite the error handling code with LOG.
'* 1.7		9/7/2022	Modify CacheQuery, add mCacheQuery
'* 1.8		10/7/2022	Add XmlCacheQuery, modify mCacheQuery,CacheQuery
'* 1.9		26/7/2022	Modify Imports, modify KeyName
'* 1.10		3/8/2022	Modify mCacheQuery
'* 1.11		4/8/2022	Modify mCacheQuery
'* 1.12		5/8/2022	Modify Property
'* 1.13		5/9/2022	Modify DebugStr
'* 1.15		5/6/2024	Modify mCacheQuery
'* 1.16     28/7/2024   Modify PigStepLog to StruStepLog
'**********************************
Imports System.Data
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
#Else
Imports Microsoft.Data.SqlClient
#End If
Imports PigToolsLiteLib


''' <summary>
''' SqlCommand for SQL Server StoredProcedure
''' </summary>
Public Class CmdSQLSrvSp
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1." & "16" & "." & "32"
    Private moSqlCommand As SqlCommand

    Public Sub New(SpName As String)
        MyBase.New(CLS_VERSION)
        Dim LOG As New StruStepLog : LOG.SubName = "New"
        Try
            mstrSpName = SpName
            moSqlCommand = New SqlCommand
            With moSqlCommand
                .CommandType = CommandType.StoredProcedure
                .CommandText = SpName
                LOG.StepName = "New SqlParameter(RETURN_VALUE)"
                Dim oSqlParameter As New SqlParameter("RETURN_VALUE", SqlDbType.Int)
                oSqlParameter.Direction = ParameterDirection.ReturnValue
                LOG.StepName = "Add(RETURN_VALUE)"
                .Parameters.Add(oSqlParameter)
                oSqlParameter = Nothing
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    ''' <summary>
    ''' Stored Procedure Name
    ''' </summary>
    Private mstrSpName As String
    Public Property SpName() As String
        Get
            Return mstrSpName
        End Get
        Set(value As String)
            mstrSpName = value
        End Set
    End Property

    ''' <summary>
    ''' Stored Procedure return value
    ''' </summary>
    Private mstrReturnValue As String
    Public ReadOnly Property ReturnValue() As Integer
        Get
            Return mstrReturnValue
        End Get
    End Property


    ''' <summary>
    ''' Execute stored procedures|执行存储过程
    ''' </summary>
    ''' <returns>Return result set|返回结果集</returns>
    Public Function Execute() As Recordset
        Dim LOG As New StruStepLog : LOG.SubName = "Execute"
        Try
            LOG.StepName = "ExecuteReader"
            Dim oSqlDataReader As SqlDataReader = moSqlCommand.ExecuteReader()
            LOG.StepName = "New Recordset"
            Execute = New Recordset(oSqlDataReader)
            If Execute.LastErr <> "" Then Throw New Exception(Execute.LastErr)
            Me.RecordsAffected = -1
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Me.RecordsAffected = -1
            Return Nothing
        End Try
    End Function



    Public Property ParaValue(ParaName As String) As Object
        Get
            Try
                ParaValue = moSqlCommand.Parameters.Item(ParaName).Value
                Me.ClearErr()
            Catch ex As Exception
                Me.SetSubErrInf("ParaValue.Get", ex)
                Return Nothing
            End Try
        End Get
        Set(value As Object)
            Try
                moSqlCommand.Parameters(ParaName).Value = value
                Me.ClearErr()
            Catch ex As Exception
                Me.SetSubErrInf("ParaValue.Set", ex)
            End Try
        End Set
    End Property

    Public Property ActiveConnection() As SqlConnection
        Get
            Try
                Return moSqlCommand.Connection
            Catch ex As Exception
                Me.SetSubErrInf("ActiveConnection.Get", ex)
                Return Nothing
            End Try
        End Get
        Set(value As SqlConnection)
            Try
                moSqlCommand.Connection = value
                Me.ClearErr()
            Catch ex As Exception
                Me.SetSubErrInf("ActiveConnection.Set", ex)
            End Try
        End Set
    End Property

    Public Sub AddPara(ParaName As String, DataType As SqlDbType)
        Dim LOG As New StruStepLog : LOG.SubName = "AddPara"
        Try
            If moSqlCommand.Parameters.IndexOf(ParaName) >= 0 Then Throw New Exception("ParaName already exists.")
            LOG.StepName = "Parameters.Add"
            moSqlCommand.Parameters.Add(ParaName, DataType)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    Public Sub AddPara(ParaName As String, DataType As SqlDbType, Size As Long)
        Dim LOG As New StruStepLog : LOG.SubName = "AddPara"
        Try
            If moSqlCommand.Parameters.IndexOf(ParaName) >= 0 Then Throw New Exception("ParaName already exists.")
            LOG.StepName = "Parameters.Add"
            moSqlCommand.Parameters.Add(ParaName, DataType, Size)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    ''' <summary>
    ''' Returns debugging information for executing SQL statements
    ''' </summary>
    Public ReadOnly Property DebugStr() As String
        Get
            Dim LOG As New StruStepLog : LOG.SubName = "DebugStr"
            Try
                LOG.StepName = "SpName"
                Dim strDebugStr As String = "EXEC " & Me.SpName & " "
                Dim bolIsBegin As Boolean = False
                If Not moSqlCommand.Parameters Is Nothing Then
                    For Each oSqlParameter As SqlParameter In moSqlCommand.Parameters
                        With oSqlParameter
                            If .Direction <> ParameterDirection.ReturnValue And Not .Value Is Nothing Then
                                LOG.StepName = "Parameters(" & .ParameterName & ")"
                                If bolIsBegin = True Then
                                    strDebugStr &= " , "
                                Else
                                    bolIsBegin = True
                                End If
                                strDebugStr &= .ParameterName & "="
                                Select Case GetDataCategoryBySqlDbType(.SqlDbType)
                                    Case Field.EnumDataCategory.BooleanValue
                                        strDebugStr &= CStr(.Value)
                                    Case Field.EnumDataCategory.DateValue
                                        strDebugStr &= mSQLStr(Format(.Value, "yyyy-MM-dd HH:mm:ss.fff"))
                                    Case Field.EnumDataCategory.IntValue, Field.EnumDataCategory.DecValue
                                        strDebugStr &= CStr(.Value)
                                    Case Field.EnumDataCategory.StrValue
                                        strDebugStr &= mSQLStr(.Value.ToString)
                                    Case Field.EnumDataCategory.OtherValue
                                        strDebugStr &= mSQLStr(.Value.ToString)
                                End Select
                            End If
                        End With
                    Next
                End If
                Return strDebugStr
            Catch ex As Exception
                Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
                Return ""
            End Try
        End Get
    End Property
    Private Function mSQLStr(SrcValue As String, Optional IsNotNull As Boolean = False) As String
        SrcValue = Replace(SrcValue, "'", "''")
        If UCase(SrcValue) = "NULL" And IsNotNull = False Then
            mSQLStr = "NULL"
        Else
            mSQLStr = "'" & SrcValue & "'"
        End If
    End Function

    '''' <summary>
    '''' Records Affected by the execution of the Stored Procedure
    '''' </summary>
    Private mlngRecordsAffected As Long
    Public Property RecordsAffected As Long
        Get
            Return mlngRecordsAffected
        End Get
        Friend Set(value As Long)
            mlngRecordsAffected = value
        End Set
    End Property

    ''' <summary>
    ''' Executing stored procedures that do not return results|执行没有返回结果的存储过程
    ''' </summary>
    ''' <returns>Execution result: OK indicates success, and others are error messages|执行结果，OK表示成功，其他为错误信息</returns>
    Public Function ExecuteNonQuery() As String
        Const SUB_NAME As String = "ExecuteNonQuery"
        Try
            Me.RecordsAffected = moSqlCommand.ExecuteNonQuery
            Return "OK"
        Catch ex As Exception
            Me.RecordsAffected = -1
            Me.PrintDebugLog(SUB_NAME, "Catch ex As Exception", Me.DebugStr)
            Return Me.GetSubErrInf(SUB_NAME, ex)
        End Try
    End Function

    ''' <summary>
    ''' 用于缓存的键值名称|The name of the key value used for caching
    ''' </summary>
    ''' <param name="HeadPartName">键值名称前缀部分|Prefix part of key name</param>
    ''' <returns></returns>
    Public ReadOnly Property KeyName(Optional HeadPartName As String = "") As String
        Get
            Try
                Dim oPigMD5 As New PigMD5(Me.DebugStr, PigMD5.enmTextType.UTF8)
                KeyName = oPigMD5.PigMD5
                If HeadPartName <> "" Then KeyName = HeadPartName & "." & KeyName
                oPigMD5 = Nothing
            Catch ex As Exception
                Me.SetSubErrInf("KeyName", ex)
                Return ""
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Query with cache, the returned result is a JSON array|带缓存的查询，返回结果为JSon数组
    ''' </summary>
    ''' <param name="ConnSQLSrv">Database Connection|数据库连接</param>
    ''' <param name="CacheTime">Cache time|缓存时间</param>
    ''' <param name="HitCache">Hit Cache Level|命中缓存级别</param>
    ''' <returns>返回的JSON数组</returns>
    Public Function CacheQuery(ByRef ConnSQLSrv As ConnSQLSrv, Optional CacheTime As Integer = 120, Optional ByRef HitCache As ConnSQLSrv.HitCacheEnum = ConnSQLSrv.HitCacheEnum.Null) As String
        Try
            CacheQuery = ""
            Dim strRet As String = Me.mCacheQuery(ConnSQLSrv, ConnSQLSrv.CacheQueryResTypeEnum.JSon, CacheQuery,, CacheTime, HitCache)
        Catch ex As Exception
            CacheQuery = ""
            Me.SetSubErrInf("CacheQuery", ex)
        End Try
    End Function

    ''' <summary>
    ''' CacheQuery with XML output|输出结果为XML文本的CacheQuery
    ''' </summary>
    ''' <param name="ConnSQLSrv">Database Connection|数据库连接</param>
    ''' <param name="OutXmlStr">XML text of the output result|输出结果的XML文本</param>
    ''' <param name="CacheTime">Cache time|缓存时间</param>
    ''' <param name="HitCache">Hit Cache Level|命中缓存级别</param>
    ''' <returns>Execution result: OK indicates success, and others are error messages|执行结果，OK表示成功，其他为错误信息</returns>
    Public Function XmlCacheQuery(ByRef ConnSQLSrv As ConnSQLSrv, ByRef OutXmlStr As String, Optional CacheTime As Integer = 120, Optional ByRef HitCache As ConnSQLSrv.HitCacheEnum = ConnSQLSrv.HitCacheEnum.Null) As String
        Return Me.mCacheQuery(ConnSQLSrv, ConnSQLSrv.CacheQueryResTypeEnum.XmlOutStr, OutXmlStr,, CacheTime, HitCache)
    End Function

    ''' <summary>
    ''' The output result is CacheQuery of XmlRS object|输出结果为XmlRS对象的CacheQuery
    ''' </summary>
    ''' <param name="ConnSQLSrv">Database Connection|数据库连接</param>
    ''' <param name="OutRS">XML text of the output result|输出结果的XML文本</param>
    ''' <param name="CacheTime">Cache time|缓存时间</param>
    ''' <param name="HitCache">Hit Cache Level|命中缓存级别</param>
    ''' <returns>Execution result: OK indicates success, and others are error messages|执行结果，OK表示成功，其他为错误信息</returns>
    Public Function XmlCacheQuery(ByRef ConnSQLSrv As ConnSQLSrv, ByRef OutRS As XmlRS, Optional CacheTime As Integer = 120, Optional ByRef HitCache As ConnSQLSrv.HitCacheEnum = ConnSQLSrv.HitCacheEnum.Null) As String
        Return Me.mCacheQuery(ConnSQLSrv, ConnSQLSrv.CacheQueryResTypeEnum.XmlOutRS,, OutRS, CacheTime, HitCache)
    End Function

    Private Function mCacheQuery(ByRef ConnSQLSrv As ConnSQLSrv, ResType As ConnSQLSrv.CacheQueryResTypeEnum, Optional ByRef OutStr As String = "", Optional ByRef OutRS As XmlRS = Nothing, Optional CacheTime As Integer = 60, Optional ByRef IsHitCache As ConnSQLSrv.HitCacheEnum = ConnSQLSrv.HitCacheEnum.List) As String
        Dim LOG As New StruStepLog : LOG.SubName = "mCacheQuery"
        Try
            With ConnSQLSrv
                Dim bolIsExec As Boolean = False, pbValue As New PigBytes
                If .PigKeyValue Is Nothing Then Throw New Exception("Not InitPigKeyValue")
                Dim strKeyName As String = Me.KeyName
                LOG.StepName = "GetKeyValue"
                LOG.Ret = .PigKeyValue.GetKeyValue(strKeyName, pbValue.Main, CacheTime, IsHitCache)
                If LOG.Ret <> "OK" Then
                    bolIsExec = True
                ElseIf pbValue.Main Is Nothing Then
                    bolIsExec = True
                ElseIf pbValue.Main.Length <= 0 Then
                    bolIsExec = True
                End If
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
                If bolIsExec = False Then
                    LOG.StepName = "UnCompress"
                    LOG.Ret = pbValue.UnCompress
                    If LOG.Ret <> "OK" Then
                        bolIsExec = True
                    End If
                End If
#End If
                If bolIsExec = True Then
                    If Me.ActiveConnection Is Nothing Then
                        LOG.StepName = "Set ActiveConnection"
                        Me.ActiveConnection = ConnSQLSrv.Connection
                    End If
                    Dim rsAny As Recordset
                    LOG.StepName = "Execute"
                    rsAny = Me.Execute
                    If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
                    Select Case ResType
                        Case ConnSQLSrv.CacheQueryResTypeEnum.JSon
                            LOG.StepName = "AllRecordset2JSon"
                            OutStr = rsAny.AllRecordset2JSon
                            If rsAny.LastErr <> "" Then Throw New Exception(rsAny.LastErr)
                        Case ConnSQLSrv.CacheQueryResTypeEnum.XmlOutStr
                            LOG.StepName = "AllRecordset2Xml(OutStr)"
                            LOG.Ret = rsAny.AllRecordset2Xml(OutStr)
                            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                        Case ConnSQLSrv.CacheQueryResTypeEnum.XmlOutRS
                            LOG.StepName = "AllRecordset2Xml(OutRS)"
                            LOG.Ret = rsAny.AllRecordset2Xml(OutRS)
                            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                            OutStr = OutRS.PigXml.XmlDocument.InnerXml
                        Case Else
                            Throw New Exception("Invalid ResType is " & ResType.ToString)
                    End Select
                    LOG.StepName = "New PigText(OutStr)"
                    Dim ptText As New PigText(OutStr, PigText.enmTextType.UTF8)
                    If ptText.LastErr <> "" Then Throw New Exception(ptText.LastErr)
                    LOG.StepName = "SaveKeyValue"
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
                    LOG.Ret = .PigKeyValue.SaveKeyValue(KeyName, ptText.CompressTextBytes)
#Else
                    LOG.Ret = .PigKeyValue.SaveKeyValue(KeyName, ptText.TextBytes)
#End If
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                Else
                    LOG.StepName = "New PigText(pbValue.Main)"
                    Dim ptText As New PigText(pbValue.Main)
                    If ptText.LastErr <> "" Then Throw New Exception(ptText.LastErr)
                    Select Case ResType
                        Case ConnSQLSrv.CacheQueryResTypeEnum.JSon, ConnSQLSrv.CacheQueryResTypeEnum.XmlOutStr
                            LOG.StepName = "oPigKeyValue.StrValue"
                            OutStr = ptText.Text
                        Case ConnSQLSrv.CacheQueryResTypeEnum.XmlOutRS
                            LOG.StepName = "New XmlRS"
                            OutRS = New XmlRS(ptText.Text)
                            If OutRS Is Nothing Then Throw New Exception("OutRS Is Nothing")
                            If OutRS.LastErr <> "" Then
                                LOG.AddStepNameInf(ptText.Text)
                                Throw New Exception(OutRS.LastErr)
                            End If
                        Case Else
                            Throw New Exception("Invalid ResType is " & ResType.ToString)
                    End Select
                End If
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


End Class
