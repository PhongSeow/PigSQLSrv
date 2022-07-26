'**********************************
'* Name: XmlRS
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 用于 Recordset.AllRecordset2Xml 返回缓存数据的结果集处理|For recordset Allrecordset2xml returns the result set processing of cached data.
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 10/7/2021
'* 1.1 11/7/2022 Modify New
'* 1.2	26/7/2022	Modify Imports
'**********************************
#If NETFRAMEWORK Then
Imports PigToolsWinLib
#Else
Imports PigToolsLiteLib
#End If

Public Class XmlRS
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.2.2"
    Public ReadOnly Property PigXml As PigXml
    Public Sub New(XmlStr As String, Optional IsChgCtrlChar As Boolean = True)
        MyBase.New(CLS_VERSION)
        Dim LOG As New PigStepLog("New")
        Try
            Me.PigXml = New PigXml(False, IsChgCtrlChar)
            LOG.StepName = "SetMainXml"
            Me.PigXml.SetMainXml(XmlStr)
            If Me.PigXml.LastErr <> "" Then Throw New Exception(Me.PigXml.LastErr)
            LOG.StepName = "InitXmlDocument"
            LOG.Ret = Me.PigXml.InitXmlDocument()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    Public ReadOnly Property TotalRS As Integer
        Get
            Try
                Return Me.PigXml.XmlDocGetInt("XmlRS.TotalRS")
            Catch ex As Exception
                Me.SetSubErrInf("TotalRS", ex)
                Return -1
            End Try
        End Get
    End Property

    Public ReadOnly Property TotalRows(RSNo As Integer) As Integer
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".TotalRows"
                Return Me.PigXml.XmlDocGetInt(strXmlKey, True)
            Catch ex As Exception
                Me.SetSubErrInf("TotalRows", ex)
                Return -1
            End Try
        End Get
    End Property

    Public ReadOnly Property TypeName(RSNo As Integer, Index As Integer) As String
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".ColInf.Col" & Index.ToString & ".TypeName"
                Return Me.PigXml.XmlDocGetStr(strXmlKey, True)
            Catch ex As Exception
                Me.SetSubErrInf("TypeName", ex)
                Return ""
            End Try
        End Get
    End Property

    Private Function mGetColNameByIndex(RSNo As Integer, Index As Integer) As String
        Try
            Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".ColInf.Col" & Index.ToString
            Return Me.PigXml.XmlDocGetStr(strXmlKey)
        Catch ex As Exception
            Me.SetSubErrInf("mGetColNameByIndex", ex)
            Return ""
        End Try
    End Function

    Private Function mGetIndexByColName(RSNo As Integer, ColName As String) As Integer
        Try
            Dim intTotalCols As Integer = Me.TotalCols(RSNo)
            mGetIndexByColName = -1
            For i = 0 To intTotalCols - 1
                Dim strColName As String = Me.ColName(RSNo, i)
                If ColName = strColName Then
                    mGetIndexByColName = i
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("mGetIndexByColName", ex)
            Return -1
        End Try
    End Function
    Public ReadOnly Property DataCategory(RSNo As Integer, Index As Integer) As String
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".ColInf.Col" & Index.ToString & ".DataCategory"
                Return Me.PigXml.XmlDocGetStr(strXmlKey, True)
            Catch ex As Exception
                Me.SetSubErrInf("DataCategory", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property ColName(RSNo As Integer, Index As Integer) As String
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".ColInf.Col" & Index.ToString
                Return Me.PigXml.XmlDocGetStr(strXmlKey)
            Catch ex As Exception
                Me.SetSubErrInf("ColName", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property StrValue(RSNo As Integer, RowNo As Integer, Index As Integer) As String
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".Rows.Row" & RowNo.ToString & ".Col" & Index.ToString
                Return Me.PigXml.XmlDocGetStr(strXmlKey)
            Catch ex As Exception
                Me.SetSubErrInf("StrValue", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property StrValue(RSNo As Integer, RowNo As Integer, ColName As String) As String
        Get
            Try
                Dim intIndex As Integer = Me.mGetIndexByColName(RSNo, ColName)
                Return Me.StrValue(RSNo, RowNo, intIndex)
            Catch ex As Exception
                Me.SetSubErrInf("StrValue", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property BoolValue(RSNo As Integer, RowNo As Integer, Index As Integer) As Boolean
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".Rows.Row" & RowNo.ToString & ".Col" & Index.ToString
                Return Me.PigXml.XmlDocGetBool(strXmlKey)
            Catch ex As Exception
                Me.SetSubErrInf("BoolValue", ex)
                Return False
            End Try
        End Get
    End Property

    Public ReadOnly Property BoolValue(RSNo As Integer, RowNo As Integer, ColName As String) As Boolean
        Get
            Try
                Dim intIndex As Integer = Me.mGetIndexByColName(RSNo, ColName)
                Return Me.BoolValue(RSNo, RowNo, intIndex)
            Catch ex As Exception
                Me.SetSubErrInf("BoolValue", ex)
                Return False
            End Try
        End Get
    End Property

    Public ReadOnly Property BoolValueEmpTrue(RSNo As Integer, RowNo As Integer, Index As Integer) As Boolean
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".Rows.Row" & RowNo.ToString & ".Col" & Index.ToString
                Return Me.PigXml.XmlDocGetBoolEmpTrue(strXmlKey)
            Catch ex As Exception
                Me.SetSubErrInf("BoolValueEmpTrue", ex)
                Return True
            End Try
        End Get
    End Property

    Public ReadOnly Property BoolValueEmpTrue(RSNo As Integer, RowNo As Integer, ColName As String) As Boolean
        Get
            Try
                Dim intIndex As Integer = Me.mGetIndexByColName(RSNo, ColName)
                Return Me.BoolValueEmpTrue(RSNo, RowNo, intIndex)
            Catch ex As Exception
                Me.SetSubErrInf("BoolValueEmpTrue", ex)
                Return False
            End Try
        End Get
    End Property

    Public ReadOnly Property DateValue(RSNo As Integer, RowNo As Integer, Index As Integer) As DateTime
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".Rows.Row" & RowNo.ToString & ".Col" & Index.ToString
                Return Me.PigXml.XmlDocGetDate(strXmlKey)
            Catch ex As Exception
                Me.SetSubErrInf("DateValue", ex)
                Return #1/1/1900#
            End Try
        End Get
    End Property

    Public ReadOnly Property DateValue(RSNo As Integer, RowNo As Integer, ColName As String) As DateTime
        Get
            Try
                Dim intIndex As Integer = Me.mGetIndexByColName(RSNo, ColName)
                Return Me.DateValue(RSNo, RowNo, intIndex)
            Catch ex As Exception
                Me.SetSubErrInf("DateValue", ex)
                Return #1/1/1900#
            End Try
        End Get
    End Property

    Public ReadOnly Property IntValue(RSNo As Integer, RowNo As Integer, Index As Integer) As Integer
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".Rows.Row" & RowNo.ToString & ".Col" & Index.ToString
                Return Me.PigXml.XmlGetInt(strXmlKey)
            Catch ex As Exception
                Me.SetSubErrInf("IntValue", ex)
                Return 0
            End Try
        End Get
    End Property

    Public ReadOnly Property IntValue(RSNo As Integer, RowNo As Integer, ColName As String) As Integer
        Get
            Try
                Dim intIndex As Integer = Me.mGetIndexByColName(RSNo, ColName)
                Return Me.IntValue(RSNo, RowNo, intIndex)
            Catch ex As Exception
                Me.SetSubErrInf("IntValue", ex)
                Return 0
            End Try
        End Get
    End Property

    Public ReadOnly Property LongValue(RSNo As Integer, RowNo As Integer, Index As Integer) As Long
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".Rows.Row" & RowNo.ToString & ".Col" & Index.ToString
                Return Me.PigXml.XmlDocGetLong(strXmlKey)
            Catch ex As Exception
                Me.SetSubErrInf("LongValue", ex)
                Return 0
            End Try
        End Get
    End Property

    Public ReadOnly Property LongValue(RSNo As Integer, RowNo As Integer, ColName As String) As Long
        Get
            Try
                Dim intIndex As Integer = Me.mGetIndexByColName(RSNo, ColName)
                Return Me.LongValue(RSNo, RowNo, intIndex)
            Catch ex As Exception
                Me.SetSubErrInf("LongValue", ex)
                Return 0
            End Try
        End Get
    End Property

    Public ReadOnly Property DecValue(RSNo As Integer, RowNo As Integer, Index As Integer) As Decimal
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".Rows.Row" & RowNo.ToString & ".Col" & Index.ToString
                Return Me.PigXml.XmlDocGetDec(strXmlKey)
            Catch ex As Exception
                Me.SetSubErrInf("DecValue", ex)
                Return 0
            End Try
        End Get
    End Property

    Public ReadOnly Property DecValue(RSNo As Integer, RowNo As Integer, ColName As String) As Decimal
        Get
            Try
                Dim intIndex As Integer = Me.mGetIndexByColName(RSNo, ColName)
                Return Me.DecValue(RSNo, RowNo, intIndex)
            Catch ex As Exception
                Me.SetSubErrInf("DecValue", ex)
                Return 0
            End Try
        End Get
    End Property


    Public ReadOnly Property TotalCols(RSNo As Integer) As Integer
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".ColInf.TotalCols"
                Return Me.PigXml.XmlDocGetInt(strXmlKey, True)
            Catch ex As Exception
                Me.SetSubErrInf("TotalCols", ex)
                Return -1
            End Try
        End Get
    End Property

    Public ReadOnly Property IsEOF(RSNo As Integer) As Boolean
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".IsEOF"
                Return Me.PigXml.XmlDocGetInt(strXmlKey, True)
            Catch ex As Exception
                Me.SetSubErrInf("IsEOF", ex)
                Return True
            End Try
        End Get
    End Property

End Class
