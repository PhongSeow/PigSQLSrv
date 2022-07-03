'**********************************
'* Name: Fields
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Field 的集合类|Collection class of field
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 6/6/2021
'* 1.0.2	21/7/2021	Modify Add
'* 1.0.3	28/7/2021	Modify Add
'* 1.1	8/6/2022	Add IsItemExists
'* 1.2	2/7/2022	Use PigBaseLocal
'************************************

Public Class Fields
    Inherits PigBaseLocal
    Implements IEnumerable(Of Field)
    Private Const CLS_VERSION As String = "1.2.1"

    Private moList As New List(Of Field)

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Try
                Return moList.Count
            Catch ex As Exception
                Me.SetSubErrInf("Count", ex)
                Return -1
            End Try
        End Get
    End Property
    Public Function GetEnumerator() As IEnumerator(Of Field) Implements IEnumerable(Of Field).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As Field
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(Name As String) As Field
        Get
            Try
                Item = Nothing
                For Each oField As Field In moList
                    If oField.Name = Name Then
                        Item = oField
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.Name", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Sub Add(NewItem As Field)
        Try
            moList.Add(NewItem)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Add.NewItem", ex)
        End Try
    End Sub

    Public Function Add(Name As String, TypeName As String, FieldTypeName As String, Index As Long) As Field
        Dim strStepName As String = ""
        Try
            strStepName = "New Field"
            Dim oField As New Field(Name, TypeName, FieldTypeName, Index)
            If oField.LastErr <> "" Then Throw New Exception(oField.LastErr)
            strStepName = "Add"
            moList.Add(oField)
            Add = oField
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Add.Text", strStepName, ex)
            Return Nothing
        End Try
    End Function


    Public Sub Remove(Name As String)
        Dim strStepName As String = ""
        Try
            strStepName = "For Each"
            For Each oField As Field In moList
                If oField.Name = Name Then
                    strStepName = "Remove"
                    moList.Remove(oField)
                    Exit For
                End If
            Next
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Remove.Name", strStepName, ex)
        End Try
    End Sub

    Public Sub Remove(Index As Integer)
        Try
            moList.RemoveAt(Index)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Remove.Name", ex)
        End Try
    End Sub

    Public Function IsItemExists(Name As String) As Boolean
        Try
            IsItemExists = False
            For Each oField As Field In moList
                If oField.Name = Name Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

End Class
