'**********************************
'* Name: Field
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Similar to ObjAdoDBLib.RecordSet
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.4
'* Create Time: 5/6/2021
'* 1.0.2	6/6/2021	Modify New
'* 1.0.3	21/7/2021	Modify New,DataCategory
'* 1.0.4	28/7/2021	Modify New,DataCategory, add FieldTypeName
'**********************************
Public Class Field
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.4"


	Public Enum DataCategoryEnum
		OtherValue = 0
		StrValue = 10
		IntValue = 20
		DecValue = 30
		BooleanValue = 40
		DateValue = 50
	End Enum


	Public Sub New(Name As String, TypeName As String, FieldTypeName As String, Index As Long)
		MyBase.New(CLS_VERSION)
		Me.Name = Name
		Me.Index = Index
		Me.TypeName = TypeName
		Me.FieldTypeName = FieldTypeName
	End Sub

	Public ReadOnly Property Name As String
	Public ReadOnly Property Index As Long
	Public ReadOnly Property TypeName As String
	Public ReadOnly Property FieldTypeName As String

	Public ReadOnly Property DataCategory() As DataCategoryEnum
		Get
			Try
				Select Case Me.FieldTypeName
					Case "String", "Guid"
						DataCategory = DataCategoryEnum.StrValue
					Case "Int64", "Boolean", "Int32", "Int16"
						DataCategory = DataCategoryEnum.IntValue
					Case "Decimal", "Double", "Single"
						DataCategory = DataCategoryEnum.DecValue
					Case "DateTime", ""
						DataCategory = DataCategoryEnum.DateValue
					Case Else
						DataCategory = DataCategoryEnum.OtherValue
				End Select
			Catch ex As Exception
				Me.SetSubErrInf("DataCategory.Get", ex)
				Return DataCategoryEnum.OtherValue
			End Try
		End Get
	End Property

	Private mintSqlDbType As SqlDbType
	Public ReadOnly Property Type() As SqlDbType
		Get
			Return mintSqlDbType
		End Get
	End Property


	Public ReadOnly Property DecValue() As Decimal
		Get
			Try
				Return CDec(moValue)
			Catch ex As Exception
				Me.SetSubErrInf("DecValue.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property DateValue() As DateTime
		Get
			Try
				Return CDate(moValue)
			Catch ex As Exception
				Me.SetSubErrInf("DateValue.Get", ex)
				Return DateTime.MinValue
			End Try
		End Get
	End Property

	Public ReadOnly Property StrValue() As String
		Get
			Try
				Return CStr(moValue)
			Catch ex As Exception
				Me.SetSubErrInf("StrValue.Get", ex)
				Return ""
			End Try
		End Get
	End Property

	Public ReadOnly Property LngValue() As Long
		Get
			Try
				Return CLng(moValue)
			Catch ex As Exception
				Me.SetSubErrInf("LngValue.Get", ex)
				Return 0
			End Try
		End Get
	End Property

	Public ReadOnly Property IntValue() As Integer
		Get
			Try
				Return CInt(moValue)
			Catch ex As Exception
				Me.SetSubErrInf("IntValue.Get", ex)
				Return 0
			End Try
		End Get
	End Property

	Public ReadOnly Property BooleanValue() As Boolean
		Get
			Try
				Return CBool(moValue)
			Catch ex As Exception
				Me.SetSubErrInf("BooleanValue.Get", ex)
				Return False
			End Try
		End Get
	End Property

	Friend ReadOnly Property ValueForJSon() As Object
		Get
			Try
				Select Case Me.DataCategory
					Case DataCategoryEnum.BooleanValue
						ValueForJSon = Me.BooleanValue
					Case DataCategoryEnum.DateValue
						ValueForJSon = Me.DateValue
					Case DataCategoryEnum.DecValue
						ValueForJSon = Me.DecValue
					Case DataCategoryEnum.IntValue
						ValueForJSon = Me.IntValue
					Case DataCategoryEnum.OtherValue
						ValueForJSon = Me.StrValue
					Case DataCategoryEnum.StrValue
						ValueForJSon = Me.StrValue
					Case Else
						ValueForJSon = Me.StrValue
				End Select
			Catch ex As Exception
				Me.SetSubErrInf("ValueForJSon.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property

	Private moValue As Object
	Public Property Value() As Object
		Get
			Try
				Return moValue
			Catch ex As Exception
				Me.SetSubErrInf("Value.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Object)
			Try
				moValue = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Value.Set", ex)
			End Try
		End Set
	End Property




End Class
