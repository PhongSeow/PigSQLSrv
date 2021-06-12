'**********************************
'* Name: CmdSQLSrvSp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: SqlCommand for SQL Server StoredProcedure
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.7
'* Create Time: 17/4/2021
'* 1.0.2	18/4/2021	Modify ActiveConnection
'* 1.0.3	24/4/2021	Add mAdoDataType
'* 1.0.4	25/4/2021	Modify New
'* 1.0.5	28/4/2021	Add ActiveConnection,AddPara,ParaValue,Execute
'* 1.0.6	16/5/2021	SQLSrvDataTypeEnum move to ConnSQLSrv, Modify Execute,ParaValue,ActiveConnection
'* 1.0.7	12/6/2021	Move to PigSQLSrvLib
'**********************************
Imports System.Data
Imports System.Data.SqlClient
Public Class CmdSQLSrvSp
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.7"
	Private moSqlCommand As SqlCommand

	Public Sub New(SpName As String)
		MyBase.New(CLS_VERSION)
		Dim strStepName As String = ""
		Try
			moSqlCommand = New SqlCommand
			With moSqlCommand
				.CommandType = CommandType.StoredProcedure
				.CommandText = SpName
				strStepName = "New SqlParameter(RETURN_VALUE)"
				Dim oSqlParameter As New SqlParameter("RETURN_VALUE", SqlDbType.Int)
				oSqlParameter.Direction = ParameterDirection.ReturnValue
				strStepName = "Add(RETURN_VALUE)"
				.Parameters.Add(oSqlParameter)
				oSqlParameter = Nothing
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("New", strStepName, ex)
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
		Set(ByVal value As String)
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
	''' Records Affected by the execution of the Stored Procedure
	''' </summary>
	Private mlngRecordsAffected As Long
	Public ReadOnly Property RecordsAffected() As Long
		Get
			Return mlngRecordsAffected
		End Get
	End Property


	Public Function Execute() As Recordset
		Dim strStepName As String = ""
		Try
			Execute = New Recordset
			With Execute
				strStepName = "ExecuteReader"
				.SqlDataReader = moSqlCommand.ExecuteReader()
				mlngRecordsAffected = .SqlDataReader.RecordsAffected
				strStepName = "New Fields"
				.Fields = New Fields
				For i = 0 To .SqlDataReader.FieldCount - 1
					strStepName = "Fields.Add（" & i & ")"
					.Fields.Add(.SqlDataReader.GetName(i), i)
					If .Fields.LastErr <> "" Then Throw New Exception(.Fields.LastErr)
				Next
				If .SqlDataReader.HasRows = True Then
					strStepName = "MoveNext"
					.MoveNext()
					If .LastErr <> "" Then Throw New Exception(.LastErr)
				End If
			End With
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


End Class
