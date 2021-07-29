'**********************************
'* Name: Recordset
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Similar to ObjAdoDBLib.RecordSet
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.6
'* Create Time: 5/6/2021
'* 1.0.2	6/6/2021	Modify EOF,Fields,MoveNext
'* 1.0.3	21/6/2021	Add Finalize,Close
'* 1.0.4	2/7/2021	Add IsTrimJSonValue,Row2JSon
'* 1.0.5	9/7/2021	Modify Row2JSon
'* 1.0.6	21/7/2021	Add Move,NextRecordset,Init, modify MoveNext,Finalize
'**********************************
Imports System.Data
Imports Microsoft.Data.SqlClient

Public Class Recordset
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.6"
	Private moSqlDataReader As SqlDataReader

	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub

	Public Sub New(SqlDataReader As SqlDataReader)
		MyBase.New(CLS_VERSION)
		Me.mNew(SqlDataReader)
	End Sub

	Private moFields As Fields
	Public Property Fields() As Fields
		Get
			Try
				Return moFields
			Catch ex As Exception
				Me.SetSubErrInf("Fields.Get", ex)
				Return Nothing
			End Try
		End Get
		Friend Set(value As Fields)
			Try
				moFields = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Fields.Set", ex)
			End Try
		End Set
	End Property

	Private mbolEOF As Boolean = True
	Public Property EOF() As Boolean
		Get
			Return mbolEOF
		End Get
		Friend Set(value As Boolean)
			Try
				mbolEOF = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("EOF.Set", ex)
				mbolEOF = True
			End Try
		End Set
	End Property

	Public Sub Close()
		Try
			If moSqlDataReader Is Nothing Then
				If moSqlDataReader.IsClosed = False Then
					moSqlDataReader.Close()
				End If
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Close", ex)
		End Try
	End Sub

	Public Sub Move(NumRecords As Long)
		Try
			Do While NumRecords > 0
				Me.MoveNext()
				If Me.EOF = True Then Exit Do
				NumRecords -= 1
			Loop
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Move", ex)
		End Try
	End Sub

	Public Sub MoveNext()
		Dim strStepName As String = ""
		Try
			strStepName = "Read"
			If moSqlDataReader.Read() = True Then
				For i = 0 To moSqlDataReader.FieldCount - 1
					strStepName = "GetValue(" & i.ToString & ")"
					Me.Fields.Item(i).Value = moSqlDataReader.GetValue(i)
				Next
				mbolEOF = False
			Else
				For i = 0 To moSqlDataReader.FieldCount - 1
					strStepName = "GetValue(" & i.ToString & ")"
					Me.Fields.Item(i).Value = Nothing
				Next
				mbolEOF = True
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("MoveNext", strStepName, ex)
			mbolEOF = True
		End Try
	End Sub

	''' <summary>
	''' Convert current row to JSON|当前行转换成JSON
	''' </summary>
	Public Function Row2JSon() As String
		Try
			Dim pjMain As New PigJSonLite
			With pjMain
				If Me.EOF = False Then
					For i = 0 To Me.Fields.Count - 1
						Dim oField As Field = Me.Fields.Item(i)
						Dim strName As String = oField.Name
						Dim strValue As String = oField.ValueForJSon
						If strName = "" Then strName = "Col" & (i + 1).ToString
						If Me.IsTrimJSonValue = True Then strValue = Trim(strValue)
						If i = 0 Then
							.AddEle(strName, strValue, True)
						Else
							.AddEle(strName, strValue)
						End If
					Next
					.AddSymbol(PigJSonLite.xpSymbolType.EleEndFlag)
				End If
			End With
			Row2JSon = pjMain.MainJSonStr
			pjMain = Nothing
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Row2JSon", ex)
			Return ""
		End Try
	End Function

	''' <summary>
	''' Whether to remove the space before and after the value is converted to JSON
	''' </summary>
	Private mbolIsTrimJSonValue As Boolean = True
	Public Property IsTrimJSonValue() As Boolean
		Get
			Return mbolIsTrimJSonValue
		End Get
		Set(ByVal value As Boolean)
			mbolIsTrimJSonValue = value
		End Set
	End Property

	Protected Overrides Sub Finalize()
		If Not moSqlDataReader Is Nothing Then
			moSqlDataReader.Close()
		End If
		MyBase.Finalize()
	End Sub

	Private Sub mNew(SqlDataReader As SqlDataReader)
		Dim strStepName As String = ""
		Try
			With Me
				strStepName = "ExecuteReader"
				moSqlDataReader = SqlDataReader
				mlngRecordsAffected = moSqlDataReader.RecordsAffected
				strStepName = "New Fields"
				.Fields = New Fields
				For i = 0 To moSqlDataReader.FieldCount - 1
					strStepName = "Fields.Add（" & i & ")"
					.Fields.Add(moSqlDataReader.GetName(i), moSqlDataReader.GetDataTypeName(i), moSqlDataReader.GetFieldType(i).Name, i)
					If .Fields.LastErr <> "" Then Throw New Exception(.Fields.LastErr)
				Next
				If moSqlDataReader.HasRows = True Then
					strStepName = "MoveNext"
					.MoveNext()
					If .LastErr <> "" Then Throw New Exception(.LastErr)
				End If
			End With
		Catch ex As Exception
			Me.SetSubErrInf("mNew", strStepName, ex)
		End Try
	End Sub
	Public Function NextRecordset() As Recordset
		Try
			If moSqlDataReader.NextResult() = False Then Throw New Exception("Has not next recordset")
			Dim oRecordset As New Recordset(moSqlDataReader)

			Me.ClearErr()
			Return oRecordset
		Catch ex As Exception
			Me.SetSubErrInf("NextRecordset", ex)
			Return Nothing
		End Try
	End Function

	''' <summary>
	''' Records Affected by the execution of the Stored Procedure
	''' </summary>
	Private mlngRecordsAffected As Long
	Public ReadOnly Property RecordsAffected() As Long
		Get
			Return mlngRecordsAffected
		End Get
	End Property

End Class
