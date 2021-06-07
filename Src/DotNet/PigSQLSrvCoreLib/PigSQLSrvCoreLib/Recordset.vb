﻿'**********************************
'* Name: Recordset
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Similar to ObjAdoDBLib.RecordSet
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.2
'* Create Time: 5/6/2021
'* 1.0.2	6/6/2021	Modify EOF,Fields,MoveNext
'**********************************
Imports System.Data
Imports Microsoft.Data.SqlClient
Public Class Recordset
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.2"
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


End Class
