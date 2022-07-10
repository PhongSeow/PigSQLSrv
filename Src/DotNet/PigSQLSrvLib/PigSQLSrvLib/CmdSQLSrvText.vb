'**********************************
'* Name: CmdSQLSrvText
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Command for SQL Server SQL statement Text
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.10
'* Create Time: 15/5/2021
'* 1.0.2	18/4/2021	Modify Execute,ParaValue
'* 1.0.3	17/5/2021	Modify ParaValue,ActiveConnection,Execute
'* 1.0.4	5/6/2021	Modify ActiveConnection,AddPara,Execute
'* 1.0.5	6/6/2021	Modify AddPara,Execute
'* 1.0.6	21/6/2021	Modify Execute
'* 1.0.7	17/7/2021	Add DebugStr,mSQLStr
'* 1.0.8	28/7/2021	Modify DebugStr
'* 1.0.9	1/8/2021	Modify DebugStr
'* 1.1		29/8/2021   Add support for .net core
'* 1.2		4/9/2021	Add RecordsAffected
'* 1.3		7/9/2021	Add ExecuteNonQuery
'* 1.4		19/9/2021	Modify Execute
'* 1.5		24/9/2021	Add KeyName,CacheQuery
'* 1.6		8/10/2021	Modify CacheQuery
'* 1.7		15/12/2021	Modify CacheQuery, and Rewrite the error handling code with LOG.
'* 1.8		2/7/2022	Use PigBaseLocal
'* 1.9		9/7/2022	Modify CacheQuery, add mCacheQuery
'* 1.10		10/7/2022	Add XmlCacheQuery, modify mCacheQuery,CacheQuery
'**********************************
Imports System.Data
Imports PigKeyCacheLib
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
#Else
Imports Microsoft.Data.SqlClient
#End If
Imports PigToolsLiteLib
Public Class CmdSQLSrvText
	Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.10.10"
    Public Property SQLText As String
	Private moSqlCommand As SqlCommand

	Public Sub New(SQLText As String)
		MyBase.New(CLS_VERSION)
		Dim LOG As New PigStepLog("New")
		Try
			Me.SQLText = SQLText
			moSqlCommand = New SqlCommand
			With moSqlCommand
				.CommandType = CommandType.Text
				.CommandText = SQLText
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Sub

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
		Dim LOG As New PigStepLog("AddPara")
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
		Dim LOG As New PigStepLog("AddPara")
		Try
			If moSqlCommand.Parameters.IndexOf(ParaName) >= 0 Then Throw New Exception("ParaName already exists.")
			LOG.StepName = "Parameters.Add"
			moSqlCommand.Parameters.Add(ParaName, DataType, Size)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Sub


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

	Public Function Execute() As Recordset
		Dim LOG As New PigStepLog("Execute")
		Me.RecordsAffected = -1
		Try
			LOG.StepName = "ExecuteReader"
			Dim oSqlDataReader As SqlDataReader = moSqlCommand.ExecuteReader()
			LOG.StepName = "New Recordset"
			Execute = New Recordset(oSqlDataReader)
			If Execute.LastErr <> "" Then Throw New Exception(Execute.LastErr)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Execute", ex)
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

	''' <summary>
	''' 用于缓存的键值名称|The name of the key value used for caching
	''' </summary>
	''' <param name="HeadPartName">键值名称前缀部分|Prefix part of key name</param>
	''' <returns></returns>
	Public ReadOnly Property KeyName(Optional HeadPartName As String = "") As String
		Get
			Try
				Dim oPigMD5 As New PigToolsLiteLib.PigMD5(Me.DebugStr, PigToolsLiteLib.PigMD5.enmTextType.UTF8)
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
	''' Returns debugging information for executing SQL statements
	''' </summary>
	Public ReadOnly Property DebugStr() As String
		Get
			Dim LOG As New PigStepLog("DebugStr")
			Try
				Dim strDebugStr As String = Me.SQLText & vbCrLf
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
                                        strDebugStr &= mSQLStr(.Value.ToString)
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
    ''' The cache query returns Recordset.AllRecordset2JSon. Note that for SQL statements with updated data, using the cache query may have unpredictable results.
    ''' </summary>
    ''' <returns></returns>
    Public Function CacheQuery(ByRef ConnSQLSrv As ConnSQLSrv, Optional CacheTime As Integer = 60) As String
        Try
            CacheQuery = ""
            Dim strRet As String = Me.mCacheQuery(ConnSQLSrv, ConnSQLSrv.CacheQueryResTypeEnum.JSon, CacheQuery,, CacheTime)
        Catch ex As Exception
            CacheQuery = ""
            Me.SetSubErrInf("CacheQuery", ex)
        End Try
    End Function

    Public Function XmlCacheQuery(ByRef ConnSQLSrv As ConnSQLSrv, ByRef OutXmlStr As String, Optional CacheTime As Integer = 60) As String
        Return Me.mCacheQuery(ConnSQLSrv, ConnSQLSrv.CacheQueryResTypeEnum.XmlOutStr, OutXmlStr,, CacheTime)
    End Function

    Public Function XmlCacheQuery(ByRef ConnSQLSrv As ConnSQLSrv, ByRef OutRS As XmlRS, Optional CacheTime As Integer = 60) As String
        Return Me.mCacheQuery(ConnSQLSrv, ConnSQLSrv.CacheQueryResTypeEnum.XmlOutRS,, OutRS, CacheTime)
    End Function

    Private Function mCacheQuery(ByRef ConnSQLSrv As ConnSQLSrv, ResType As ConnSQLSrv.CacheQueryResTypeEnum, Optional ByRef OutStr As String = "", Optional ByRef OutRS As XmlRS = Nothing, Optional CacheTime As Integer = 60) As String
        Dim LOG As New PigStepLog("mCacheQuery")
        Try
            With ConnSQLSrv
                If .PigKeyValueApp Is Nothing Then
                    LOG.StepName = "InitPigKeyValue"
                    .InitPigKeyValue()
                    If .LastErr <> "" Then Throw New Exception(.LastErr)
                End If
                Dim strKeyName As String = Me.KeyName
                LOG.StepName = "GetPigKeyValue"
                Dim oPigKeyValue As PigKeyValue = .PigKeyValueApp.GetPigKeyValue(strKeyName)
                If .PigKeyValueApp.LastErr <> "" Then Throw New Exception(.PigKeyValueApp.LastErr)
                Dim bolIsExec As Boolean = False
                If oPigKeyValue Is Nothing Then
                    bolIsExec = True
                ElseIf oPigKeyValue.IsExpired = True Then
                    bolIsExec = True
                End If
                If bolIsExec = True Then
                    If Me.ActiveConnection Is Nothing Then
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
                            LOG.StepName = "AllRecordset2JSon(XmlOutStr)"
                            LOG.Ret = rsAny.AllRecordset2Xml(OutStr)
                            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                        Case ConnSQLSrv.CacheQueryResTypeEnum.XmlOutRS
                            LOG.StepName = "AllRecordset2JSon(OutRS)"
                            LOG.Ret = rsAny.AllRecordset2Xml(OutRS)
                            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                        Case Else
                            Throw New Exception("Invalid ResType is " & ResType.ToString)
                    End Select
                    LOG.StepName = "New PigKeyValue"
                    Select Case ResType
                        Case ConnSQLSrv.CacheQueryResTypeEnum.XmlOutRS
                            oPigKeyValue = New PigKeyValue(strKeyName, Now.AddSeconds(CacheTime), OutRS.PigXml.MainXmlStr, PigToolsLiteLib.PigText.enmTextType.UTF8, PigKeyValue.EnmSaveType.SaveSpace)
                        Case Else
                            oPigKeyValue = New PigKeyValue(strKeyName, Now.AddSeconds(CacheTime), OutStr, PigToolsLiteLib.PigText.enmTextType.UTF8, PigKeyValue.EnmSaveType.SaveSpace)
                    End Select
                    If oPigKeyValue.LastErr <> "" Then Throw New Exception(oPigKeyValue.LastErr)
                    LOG.StepName = "PigKeyValueApp.SavePigKeyValue"
                    .PigKeyValueApp.SavePigKeyValue(oPigKeyValue)
                    If .PigKeyValueApp.LastErr <> "" Then Throw New Exception(.PigKeyValueApp.LastErr)
                End If
                mCacheQuery = oPigKeyValue.StrValue
                oPigKeyValue = Nothing
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return ""
        End Try
    End Function

End Class
