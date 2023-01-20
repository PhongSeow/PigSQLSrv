'**********************************
'* Name: XmlRS
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Xml result set|Xml结果集
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.7
'* Create Time: 10/7/2021
'* 1.1 11/7/2022 Modify New
'* 1.2	26/7/2022	Modify Imports
'* 1.3	28/7/2022	Add IsColExists
'* 1.5	5/9/2022	Modify datetime
'* 1.6	27/9/2022	Modify IntValue
'* 1.7	10/10/2022	Modify IsEOF
'**********************************
#If NETFRAMEWORK Then
Imports PigToolsWinLib
#Else
Imports PigToolsLiteLib
#End If


''' <summary>
''' Xml result set|Xml结果集
''' </summary>
Public Class XmlRS
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.7.1"
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

    ''' <summary>
    ''' Whether the specified column exists|指定的列是否存在
    ''' </summary>
    ''' <param name="RSNo"></param>
    ''' <param name="ColName">列名</param>
    ''' <returns></returns>
    Public ReadOnly Property IsColExists(RSNo As Integer, ColName As String) As Boolean
        Get
            Try
                IsColExists = False
                For i = 0 To Me.TotalCols(RSNo) - 1
                    If Me.ColName(RSNo, i) = ColName Then
                        IsColExists = True
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("IsColExists", ex)
                Return False
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Total number of result sets|结果集的总数
    ''' </summary>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Number of rows in the result set|结果集的行数
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' 列的数据类型
    ''' </summary>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="Index">Index value of the column|列的索引值</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Specify the text value of a column and a row in the result set|指定结果集某列某行的文本值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="Index">Index value of the column|列的索引值</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Specify the text value of a column and a row in the result set|指定结果集某列某行的文本值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="ColName">Column name|列名</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Whether the specified column exists|指定结果集某列某行的布尔值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="ColName">Column name|列名</param>
    ''' <returns></returns>
    Public ReadOnly Property BooleanValue(RSNo As Integer, RowNo As Integer, ColName As String) As Boolean
        Get
            Try
                Dim intIndex As Integer = Me.mGetIndexByColName(RSNo, ColName)
                Return Me.BooleanValue(RSNo, RowNo, intIndex)
            Catch ex As Exception
                Me.SetSubErrInf("BooleanValue", ex)
                Return False
            End Try
        End Get
    End Property


    ''' <summary>
    ''' Whether the specified column exists|指定结果集某列某行的布尔值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="Index">Index value of the column|列的索引值</param>
    ''' <returns></returns>
    Public ReadOnly Property BooleanValue(RSNo As Integer, RowNo As Integer, Index As Integer) As Boolean
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".Rows.Row" & RowNo.ToString & ".Col" & Index.ToString
                Return Me.PigXml.XmlDocGetBool(strXmlKey)
            Catch ex As Exception
                Me.SetSubErrInf("BooleanValue", ex)
                Return False
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Whether the specified column exists|指定结果集某列某行的布尔值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="Index">Index value of the column|列的索引值</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Whether the specified column exists|指定结果集某列某行的布尔值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="ColName">Column name|列名</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Specify the Boolean value of a column and a row in the result set. Null is TRUE|指定结果集某列某行的布尔值，空为TRUE
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="Index">Index value of the column|列的索引值</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Specify the Boolean value of a column and a row in the result set. Null is TRUE|指定结果集某列某行的布尔值，空为TRUE
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="ColName">Column name|列名</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Specify the date value of a column and a row in the result set|指定结果集某列某行的日期值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="Index">Index value of the column|列的索引值</param>
    ''' <returns></returns>
    Public ReadOnly Property DateValue(RSNo As Integer, RowNo As Integer, Index As Integer) As Date
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

    ''' <summary>
    ''' Specify the date value of a column and a row in the result set|指定结果集某列某行的日期值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="ColName">Column name|列名</param>
    ''' <returns></returns>
    Public ReadOnly Property DateValue(RSNo As Integer, RowNo As Integer, ColName As String) As Date
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

    ''' <summary>
    ''' Specify the integer value of a column and a row in the result set|指定结果集某列某行的整数值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="Index">Index value of the column|列的索引值</param>
    ''' <returns></returns>
    Public ReadOnly Property IntValue(RSNo As Integer, RowNo As Integer, Index As Integer) As Integer
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".Rows.Row" & RowNo.ToString & ".Col" & Index.ToString
                Return Me.PigXml.XmlDocGetInt(strXmlKey)
            Catch ex As Exception
                Me.SetSubErrInf("IntValue", ex)
                Return 0
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Specify the integer value of a column and a row in the result set|指定结果集某列某行的整数值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="ColName">Column name|列名</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Specify the long integer value of a column and a row in the result set|指定结果集某列某行的长整型值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="Index">Index value of the column|列的索引值</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Specify the long integer value of a column and a row in the result set|指定结果集某列某行的长整型值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="ColName">Column name|列名</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Specify the decimal value of a column and a row in the result set|指定结果集某列某行的小数值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="Index">Index value of the column|列的索引值</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' Specify the decimal value of a column and a row in the result set|指定结果集某列某行的小数值
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <param name="RowNo">Row number|行号</param>
    ''' <param name="ColName">Column name|列名</param>
    ''' <returns></returns>
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


    ''' <summary>
    ''' Number of columns in the result set|结果集的列数
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' To the last row of the current result set|是否到当前结果集的最后一行
    ''' </summary>
    ''' <param name="RSNo">Result set No|结果集号</param>
    ''' <returns></returns>
    Public ReadOnly Property IsEOF(RSNo As Integer) As Boolean
        Get
            Try
                Dim strXmlKey As String = "XmlRS.RS" & RSNo & ".IsEOF"
                Return Me.PigXml.XmlDocGetBool(strXmlKey, True)
            Catch ex As Exception
                Me.SetSubErrInf("IsEOF", ex)
                Return True
            End Try
        End Get
    End Property

End Class
