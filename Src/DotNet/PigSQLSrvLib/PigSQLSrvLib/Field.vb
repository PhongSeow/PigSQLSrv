'**********************************
'* Name: Field
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Similar to ObjAdoDBLib.RecordSet
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.8
'* Create Time: 5/6/2021
'* 1.0.2	6/6/2021	Modify New
'* 1.0.3	21/7/2021	Modify New,DataCategory
'* 1.0.4	28/7/2021	Modify New,DataCategory, add FieldTypeName
'* 1.1		29/8/2021   Add support for .net core
'* 1.2		9/6/2022    Modify EnumDataCategory,ValueForJSon
'* 1.3		24/6/2022   Rename DataCategoryEnum to EnumDataCategory
'* 1.4		2/7/2022	Use PigBaseLocal
'* 1.5		3/7/2022	Modify IntValue,LngValue,DecValue,DateValue,StrValue
'* 1.6		10/7/2022	Modify ValueForJSon, add DateFormat, add IsStrValueTrim
'* 1.7		11/7/2022	Modify DataCategory
'* 1.8		5/9/2022	Modify DateValue, add DateStrValue
'**********************************
Imports System.Data
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
#Else
Imports Microsoft.Data.SqlClient
#End If
Public Class Field
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.8.1"


    Public Enum EnumDataCategory
        OtherValue = 0
        StrValue = 10
        IntValue = 20
        DecValue = 30
        BooleanValue = 40
        DateValue = 50
        LongValue = 60
        BinaryValue = 70
    End Enum


    Public Sub New(Name As String, TypeName As String, FieldTypeName As String, Index As Long)
        MyBase.New(CLS_VERSION)
        Me.Name = Name
        Me.Index = Index
        Me.TypeName = TypeName
        Me.FieldTypeName = FieldTypeName
    End Sub

    ''' <summary>
    ''' 是否对字符串值去除前后空格|Whether to remove the space before and after the string value
    ''' </summary>
    ''' <returns></returns>
    Public Property IsStrValueTrim As Boolean = True

    Public ReadOnly Property Name As String
    Public ReadOnly Property Index As Long
    Public ReadOnly Property TypeName As String
    Public ReadOnly Property FieldTypeName As String

    Public Property DateFormat As String = "yyyy-MM-dd HH:mm:ss.fff"

    Public ReadOnly Property DataCategory() As EnumDataCategory
        Get
            Try
                Select Case Me.FieldTypeName
                    Case "String", "Guid"
                        DataCategory = EnumDataCategory.StrValue
                    Case "Int32", "Int16", "Byte"
                        DataCategory = EnumDataCategory.IntValue
                    Case "Int64"
                        DataCategory = EnumDataCategory.LongValue
                    Case "Boolean"
                        DataCategory = EnumDataCategory.BooleanValue
                    Case "Decimal", "Double", "Single"
                        DataCategory = EnumDataCategory.DecValue
                    Case "DateTime", "DateTimeOffset"
                        DataCategory = EnumDataCategory.DateValue
                    Case "Byte[]"
                        DataCategory = EnumDataCategory.BinaryValue
                    Case Else
                        DataCategory = EnumDataCategory.OtherValue
                End Select
            Catch ex As Exception
                Me.SetSubErrInf("DataCategory.Get", ex)
                Return EnumDataCategory.OtherValue
            End Try
        End Get
    End Property

    'Private mintSqlDbType As SqlDbType
    'Public ReadOnly Property Type() As SqlDbType
    '    Get
    '        Return mintSqlDbType
    '    End Get
    'End Property


    Public ReadOnly Property DecValue() As Decimal
        Get
            Try
                If IsDBNull(moValue) = True Then
                    Return 0
                ElseIf IsNumeric(moValue) = True Then
                    Return CDec(moValue)
                Else
                    Return 0
                End If
            Catch ex As Exception
                Me.SetSubErrInf("DecValue.Get", ex)
                Return 0
            End Try
        End Get
    End Property

    Public ReadOnly Property DateStrValue() As String
        Get
            Try
                If IsDBNull(moValue) = True Then
                    Return Format(#1/1/1900#, Me.DateFormat)
                ElseIf IsDate(moValue) = True Then
                    Return Format(CDate(moValue), Me.DateFormat)
                Else
                    Return Format(#1/1/1900#, Me.DateFormat)
                End If
            Catch ex As Exception
                Me.SetSubErrInf("DateValue.Get", ex)
                Return Format(#1/1/1900#, Me.DateFormat)
            End Try
        End Get
    End Property

    Public ReadOnly Property DateValue() As Date
        Get
            Try
                If IsDBNull(moValue) = True Then
                    Return #1/1/1900#
                ElseIf IsDate(moValue) = True Then
                    Return CDate(moValue)
                Else
                    Return #1/1/1900#
                End If
            Catch ex As Exception
                Me.SetSubErrInf("DateValue.Get", ex)
                Return #1/1/1900#
            End Try
        End Get
    End Property

    Public ReadOnly Property StrValue() As String
        Get
            Try
                If IsDBNull(moValue) = True Then
                    Return ""
                ElseIf Me.IsStrValueTrim = True Then
                    Return Trim(CStr(moValue))
                Else
                    Return CStr(moValue)
                End If
            Catch ex As Exception
                Me.SetSubErrInf("StrValue.Get", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property LngValue() As Long
        Get
            Try
                If IsDBNull(moValue) = True Then
                    Return 0
                ElseIf IsNumeric(moValue) = True Then
                    Return CLng(moValue)
                Else
                    Return 0
                End If
            Catch ex As Exception
                Me.SetSubErrInf("LngValue.Get", ex)
                Return 0
            End Try
        End Get
    End Property

    Public ReadOnly Property IntValue() As Integer
        Get
            Try
                If IsDBNull(moValue) = True Then
                    Return 0
                ElseIf IsNumeric(moValue) = True Then
                    Return CInt(moValue)
                Else
                    Return 0
                End If
            Catch ex As Exception
                Me.SetSubErrInf("IntValue.Get", ex)
                Return 0
            End Try
        End Get
    End Property

    Public ReadOnly Property BooleanValue() As Boolean
        Get
            Try
                If IsNumeric(moValue) = False Then
                    Select Case UCase(moValue)
                        Case "TRUE", "FALSE"
                            Return CBool(moValue)
                        Case Else
                            Return False
                    End Select
                ElseIf IsDBNull(moValue) = True Then
                    Return False
                Else
                    Return CBool(moValue)
                End If
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
                    Case EnumDataCategory.BooleanValue
                        ValueForJSon = Me.BooleanValue.ToString
                    Case EnumDataCategory.DateValue
                        ValueForJSon = Format(Me.DateValue, Me.DateFormat)
                    Case EnumDataCategory.DecValue
                        ValueForJSon = Me.DecValue.ToString
                    Case EnumDataCategory.IntValue
                        ValueForJSon = Me.IntValue.ToString
                    Case EnumDataCategory.LongValue
                        ValueForJSon = Me.LngValue.ToString
                    Case EnumDataCategory.OtherValue
                        ValueForJSon = Me.StrValue
                    Case EnumDataCategory.StrValue
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
