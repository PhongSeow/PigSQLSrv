'**********************************
'* Name: CmdSQLSrvText
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Command for SQL Server SQL statement Text
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.8
'* Create Time: 15/5/2021
'* 1.0.2	18/4/2021	Modify Execute,ParaValue
'* 1.0.3	17/5/2021	Modify ParaValue,ActiveConnection,Execute
'* 1.0.4	5/6/2021	Modify ActiveConnection,AddPara,Execute
'* 1.0.5	6/6/2021	Modify AddPara,Execute
'* 1.0.6	21/6/2021	Modify Execute
'* 1.0.7	17/7/2021	Add DebugStr,mSQLStr
'* 1.0.8	28/7/2021	Modify DebugStr
'**********************************
Imports System.Data
Imports System.Data.SqlClient
Public Class CmdSQLSrvText
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.8"
	Public Property SQLText As String
	Private moSqlCommand As SqlCommand

	Public Sub New(SQLText As String)
		MyBase.New(CLS_VERSION)
		Dim strStepName As String = ""
		Try
			Me.SQLText = SQLText
			moSqlCommand = New SqlCommand
			With moSqlCommand
				.CommandType = CommandType.Text
				.CommandText = SQLText
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("New", strStepName, ex)
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
		Dim strStepName As String = ""
		Try
			If moSqlCommand.Parameters.IndexOf(ParaName) >= 0 Then Throw New Exception("ParaName already exists.")
			strStepName = "Parameters.Add"
			moSqlCommand.Parameters.Add(ParaName, DataType)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("AddPara", strStepName, ex)
		End Try
	End Sub

	Public Sub AddPara(ParaName As String, DataType As SqlDbType, Size As Long)
		Dim strStepName As String = ""
		Try
			If moSqlCommand.Parameters.IndexOf(ParaName) >= 0 Then Throw New Exception("ParaName already exists.")
			strStepName = "Parameters.Add"
			moSqlCommand.Parameters.Add(ParaName, DataType, Size)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("AddPara", strStepName, ex)
		End Try
	End Sub


	Private Sub mGetRow()
		Try

			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mGetRow", ex)
		End Try
	End Sub


	Public Function Execute() As Recordset
		Dim strStepName As String = ""
		Try
			strStepName = "ExecuteReader"
			Dim oSqlDataReader As SqlDataReader = moSqlCommand.ExecuteReader()
			strStepName = "New Recordset"
			Execute = New Recordset(oSqlDataReader)
			If Execute.LastErr <> "" Then Throw New Exception(Execute.LastErr)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Execute", ex)
			Return Nothing
		End Try
	End Function

	'''' <summary>
	'''' Records Affected by the execution of the Stored Procedure
	'''' </summary>
	'Private mlngRecordsAffected As Long
	'Public ReadOnly Property RecordsAffected() As Long
	'	Get
	'		Return mlngRecordsAffected
	'	End Get
	'End Property

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
	''' Returns debugging information for executing SQL statements
	''' </summary>
	Public ReadOnly Property DebugStr() As String
		Get
			Dim strStepName As String = ""
			Try
				Dim strDebugStr As String = Me.SQLText & vbCrLf
				Dim bolIsBegin As Boolean = False
				If Not moSqlCommand.Parameters Is Nothing Then
					For Each oSqlParameter As SqlParameter In moSqlCommand.Parameters
						With oSqlParameter
							If .Direction <> ParameterDirection.ReturnValue Then
								strStepName = "Parameters(" & .ParameterName & ")"
								If bolIsBegin = True Then
									strDebugStr &= " , "
								Else
									bolIsBegin = True
								End If
								strDebugStr &= .ParameterName & "=" & mSQLStr(.Value.ToString)
							End If
						End With
					Next
				End If
				Return strDebugStr
			Catch ex As Exception
				Me.SetSubErrInf("DebugStr", strStepName, ex)
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

End Class
