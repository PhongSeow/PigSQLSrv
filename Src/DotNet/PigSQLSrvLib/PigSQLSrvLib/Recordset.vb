'**********************************
'* Name: Recordset
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Similar to ObjAdoDBLib.RecordSet
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.3
'* Create Time: 5/6/2021
'* 1.0.2	6/6/2021	Modify EOF,Fields,MoveNext
'* 1.0.3	21/6/2021	Add Finalize,Close
'**********************************
Imports System.Data
Imports System.Data.SqlClient
Public Class Recordset
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.3"
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

	Protected Overrides Sub Finalize()
		If Me.SqlDataReader Is Nothing Then
			Me.SqlDataReader.Close()
		End If
		MyBase.Finalize()
	End Sub
End Class
