'**********************************
'* Name: Recordset
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Similar to ObjAdoDBLib.RecordSet
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 5/6/2021
'* 1.0.2	6/6/2021	Modify EOF,Fields,MoveNext
'* 1.0.3	21/6/2021	Add Finalize,Close
'* 1.0.4	2/7/2021	Add IsTrimJSonValue,Row2JSon
'* 1.0.5	9/7/2021	Modify Row2JSon
'* 1.0.6	21/7/2021	Add Move,NextRecordset,Init, modify MoveNext,Finalize
'* 1.0.7	29/7/2021	Add Recordset2JSon,Recordset2SimpleJSonArray,MaxToJSonRows,mRecordset2JSon,AllRecordset2JSon
'* 1.1		29/8/2021   Add support for .net core
'**********************************
Imports System.Data
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
#Else
Imports Microsoft.Data.SqlClient
#End If
Public Class Recordset
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.1.2"
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
	''' Convert current recordset to JSON|当前结果集转换成JSON
	''' </summary>
	Public Function Recordset2JSon() As String
		Return Me.mRecordset2JSon()
	End Function

	''' <summary>
	''' Convert current recordset to JSON|当前结果集转换成JSON
	''' </summary>
	Public Function Recordset2JSon(TopRows As Long) As String
		Return Me.mRecordset2JSon(TopRows)
	End Function

	Private Function mRecordset2JSon(Optional TopRows As Long = -1) As String
		Dim strStepName As String = ""
		Try
			Dim intRowNo As Integer = 0
			strStepName = "New PigJSon"
			Dim pjMain As New PigJSonLite
			If pjMain.LastErr <> "" Then Throw New Exception(pjMain.LastErr)
			pjMain.AddArrayEleBegin("ROW", True)
			Do While Not Me.EOF
				If intRowNo >= Me.MaxToJSonRows Then Exit Do
				If TopRows > 0 Then
					If intRowNo >= TopRows Then Exit Do
				End If
				strStepName = "Row2JSon"
				Dim strRowJSon As String = Me.Row2JSon
				If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
				If intRowNo = 0 Then
					pjMain.AddArrayEleValue(strRowJSon, True)
				Else
					pjMain.AddArrayEleValue(strRowJSon)
				End If
				intRowNo += 1
				strStepName = "MoveNext"
				Me.MoveNext()
				If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Loop
			pjMain.AddSymbol(PigJSonLite.xpSymbolType.ArrayEndFlag)
			pjMain.AddEle("TotalRows", intRowNo)
			pjMain.AddEle("IsEOF", Me.EOF)
			pjMain.AddSymbol(PigJSonLite.xpSymbolType.EleEndFlag)
			mRecordset2JSon = pjMain.MainJSonStr
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mRecordset2JSon", strStepName, ex)
			Return ""
		End Try
	End Function

	''' <summary>
	''' Convert recordset to simple JSON array, The returned result cannot be used as a standalone JSON.|当前结果集转换成简单的JSON数组，返回结果不能作为独立的 JSon 使用。
	''' </summary>
	Public Function Recordset2SimpleJSonArray() As String
		Dim strStepName As String = ""
		Try
			Dim intRowNo As Integer = 0
			strStepName = "New PigJSonLite"
			Dim pjMain As New PigJSonLite
			If pjMain.LastErr <> "" Then Throw New Exception(pjMain.LastErr)
			Do While Not Me.EOF
				If intRowNo >= Me.MaxToJSonRows Then Exit Do
				strStepName = "Row2JSon"
				Dim strRowJSon As String = Me.Row2JSon
				If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
				If intRowNo = 0 Then
					pjMain.AddArrayEleValue(strRowJSon, True)
				Else
					pjMain.AddArrayEleValue(strRowJSon)
				End If
				intRowNo += 1
				strStepName = "MoveNext"
				Me.MoveNext()
				If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Loop
			'pjMain.AddSymbol(PigJSonLite.xpSymbolType.ArrayEndFlag)
			Recordset2SimpleJSonArray = pjMain.MainJSonStr
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Recordset2SimpleJSonArray", ex)
			Return ""
		End Try
	End Function

	''' <summary>
	''' Convert all recordset to JSON|所有结果集转换成JSON
	''' </summary>
	''' <returns></returns>
	Public Function AllRecordset2JSon() As String
		Dim strStepName As String = ""
		Try
			Dim intRSNo As Integer = 0
			strStepName = "New PigJSonLite"
			Dim pjMain As New PigJSonLite
			If pjMain.LastErr <> "" Then Throw New Exception(pjMain.LastErr)
			pjMain.AddArrayEleBegin("RS", True)
			Dim strRsJSon As String
			strStepName = "Me.Recordset2JSon"
			strRsJSon = Me.Recordset2JSon(Me.MaxToJSonRows)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			pjMain.AddArrayEleValue(strRsJSon, True)
			intRSNo = 1
			strStepName = "Me.NextRecordset"
			Dim rsParent As Recordset = Nothing
			Dim rsSub As Recordset = Me.NextRecordset
			Do While Not rsSub Is Nothing
				strStepName = "rs.Recordset2JSon"
				strRsJSon = rsSub.Recordset2JSon(Me.MaxToJSonRows)
				If rsSub.LastErr <> "" Then Throw New Exception(rsSub.LastErr)
				pjMain.AddArrayEleValue(strRsJSon)
				intRSNo += 1
				rsParent = rsSub
				strStepName = "rs.NextRecordset"
				rsSub = Nothing
				rsSub = rsParent.NextRecordset
				If rsParent.LastErr <> "" Then Exit Do
			Loop
			pjMain.AddSymbol(PigJSonLite.xpSymbolType.ArrayEndFlag)
			pjMain.AddEle("TotalRS", intRSNo)
			pjMain.AddSymbol(PigJSonLite.xpSymbolType.EleEndFlag)
			AllRecordset2JSon = pjMain.MainJSonStr
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("AllRecordset2JSon", ex)
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

	''' <summary>
	''' The maximum number of rows to convert the Recordset to JSON
	''' </summary>
	Private mlngMaxToJSonRows As Long = 1024
	Public Property MaxToJSonRows() As Long
		Get
			Return mlngMaxToJSonRows
		End Get
		Set(ByVal value As Long)
			mlngMaxToJSonRows = value
		End Set
	End Property

End Class
