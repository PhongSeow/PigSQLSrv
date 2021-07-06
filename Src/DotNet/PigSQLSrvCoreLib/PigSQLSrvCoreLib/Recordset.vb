'**********************************
'* Name: Recordset
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Similar to ObjAdoDBLib.RecordSet
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.4
'* Create Time: 5/6/2021
'* 1.0.2	6/6/2021	Modify EOF,Fields,MoveNext
'* 1.0.3	21/6/2021	Add Finalize,Close
'* 1.0.4	2/7/2021	Add IsTrimJSonValue,Row2JSon
'**********************************
Imports System.Data
Imports Microsoft.Data.SqlClient
Public Class Recordset
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.4"
	Public SqlDataReader As SqlDataReader

	Public Sub New()
		MyBase.New(CLS_VERSION)
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
			Me.SqlDataReader.Close()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Close", ex)
		End Try
	End Sub

	Public Sub MoveNext()
		Try
			If Me.SqlDataReader.Read() = True Then
				For i = 0 To Me.SqlDataReader.FieldCount - 1
					Me.Fields.Item(i).Value = Me.SqlDataReader.GetValue(i)
				Next
				mbolEOF = False
			Else
				mbolEOF = True
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("MoveNext", ex)
			mbolEOF = True
		End Try
	End Sub

	''' <summary>
	''' Convert current row to JSON|当前行转换成JSON
	''' </summary>
	Public Function Row2JSon() As String
		Try
			Dim pjMain As New PigJSon
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
					.AddSymbol(PigJSon.xpSymbolType.EleEndFlag)
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
		If Me.SqlDataReader Is Nothing Then
			Me.SqlDataReader.Close()
		End If
		MyBase.Finalize()
	End Sub
End Class
