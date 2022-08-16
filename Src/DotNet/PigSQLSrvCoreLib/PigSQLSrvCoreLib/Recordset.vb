'**********************************
'* Name: Recordset
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Similar to ObjAdoDBLib.RecordSet
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.11
'* Create Time: 5/6/2021
'* 1.0.2	6/6/2021	Modify EOF,Fields,MoveNext
'* 1.0.3	21/6/2021	Add Finalize,Close
'* 1.0.4	2/7/2021	Add IsTrimJSonValue,Row2JSon
'* 1.0.5	9/7/2021	Modify Row2JSon
'* 1.0.6	21/7/2021	Add Move,NextRecordset,Init, modify MoveNext,Finalize
'* 1.0.7	29/7/2021	Add Recordset2JSon,Recordset2SimpleJSonArray,MaxToJSonRows,mRecordset2JSon,AllRecordset2JSon
'* 1.1		29/8/2021   Add support for .net core
'* 1.2		29/8/2021   Modify Close
'* 1.3		15/12/2021	Rewrite the error handling code with LOG.
'* 1.4		2/7/2022	Use PigBaseLocal
'* 1.5		3/7/2022	Modify NextRecordset
'* 1.6		9/7/2022	Add mRecordset2Xml,Recordset2Xml,AllRecordset2Xml
'* 1.7		11/7/2022	Modify mRecordset2Xml,AllRecordset2Xml,mGetRSColInfXml
'* 1.8	    26/7/2022	Modify Imports
'* 1.9		29/7/2022	Modify Imports
'* 1.10		3/8/2022	Modify AllRecordset2Xml
'* 1.11		5/8/2022	Modify Property
'**********************************
Imports System.Data
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
#Else
Imports Microsoft.Data.SqlClient
#End If
Imports PigToolsLiteLib
Public Class Recordset
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.11.2"
    Private moSqlDataReader As SqlDataReader


    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    ''' <summary>
    ''' 是否对字符串值去除前后空格|Whether to remove the space before and after the string value
    ''' </summary>
    ''' <returns></returns>
    Public Property IsStrValueTrim As Boolean = True

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
            If Not moSqlDataReader Is Nothing Then
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
        Dim LOG As New PigStepLog("MoveNext")
        Try
            LOG.StepName = "Read"
            If moSqlDataReader.Read() = True Then
                For i = 0 To moSqlDataReader.FieldCount - 1
                    LOG.StepName = "GetValue(" & i.ToString & ")"
                    Me.Fields.Item(i).Value = moSqlDataReader.GetValue(i)
                Next
                mbolEOF = False
            Else
                For i = 0 To moSqlDataReader.FieldCount - 1
                    LOG.StepName = "GetValue(" & i.ToString & ")"
                    Me.Fields.Item(i).Value = Nothing
                Next
                mbolEOF = True
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
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
                        oField.IsStrValueTrim = Me.IsStrValueTrim
                        Dim strName As String = oField.Name
                        Dim strValue As String = oField.ValueForJSon
                        If strName = "" Then strName = "Col" & (i + 1).ToString
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

    ''' <summary>
    ''' 获取当前结果集的列信息|Get the column information of the current result set
    ''' </summary>
    ''' <param name="OutXml">输出的XML片段|Output XML fragment</param>
    ''' <returns></returns>
    Private Function mGetRSColInfXml(ByRef OutXml) As String
        Dim LOG As New PigStepLog("mGetRSColInfXml")
        Try
            LOG.StepName = "New PigXml"
            Dim oPigXml As New PigXml(False)
            If oPigXml.LastErr <> "" Then Throw New Exception(oPigXml.LastErr)
            oPigXml.AddEleLeftSign("ColInf", True)
            oPigXml.AddEleLeftAttribute("TotalCols", Me.Fields.Count.ToString)
            oPigXml.AddEleLeftSignEnd()
            For i = 0 To Me.Fields.Count - 1
                Dim strCol As String = "Col" & i.ToString
                LOG.StepName = "Add " & i.ToString
                oPigXml.AddEleLeftSign(strCol, True)
                With Me.Fields.Item(i)
                    .IsStrValueTrim = Me.IsStrValueTrim
                    oPigXml.AddEleLeftAttribute("TypeName", .TypeName)
                    oPigXml.AddEleLeftAttribute("DataCategory", .DataCategory.ToString)
                    oPigXml.AddEleLeftSignEnd()
                    Dim strName As String = .Name
                    If strName = "" Then strName = "Col" & i.ToString
                    oPigXml.AddXmlFragment(strName)
                End With
                oPigXml.AddEleRightSign(strCol)
                If oPigXml.LastErr <> "" Then Throw New Exception(oPigXml.LastErr)
            Next
            oPigXml.AddEleRightSign("ColInf")
            OutXml = oPigXml.MainXmlStr
            oPigXml = Nothing
            Return "OK"
        Catch ex As Exception
            OutXml = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function Recordset2Xml(ByRef OutXml As String, TopRows As Long) As String
        Return Me.mRecordset2Xml(OutXml, TopRows)
    End Function

    Public Function Recordset2Xml(ByRef OutXml As String) As String
        Return Me.mRecordset2Xml(OutXml)
    End Function

    Private Function mRecordset2Xml(ByRef OutXml As String, Optional TopRows As Long = -1, Optional RSNo As Integer = 0) As String
        Dim LOG As New PigStepLog("mRecordset2Xml")
        Try
            Dim intRowNo As Integer = 0
            LOG.StepName = "New PigXml"
            Dim oPigXml As New PigXml(False)
            If oPigXml.LastErr <> "" Then Throw New Exception(oPigXml.LastErr)
            Do While Not Me.EOF
                intRowNo += 1
                If intRowNo >= Me.MaxTopJSonOrXmlRows Then Exit Do
                If TopRows > 0 Then
                    If intRowNo >= TopRows Then Exit Do
                End If
                Dim strRow As String = "Row" & intRowNo.ToString
                oPigXml.AddEleLeftSign(strRow)
                For i = 0 To Me.Fields.Count - 1
                    oPigXml.AddEle("Col" & i.ToString, Me.Fields.Item(i).ValueForJSon)
                Next
                LOG.StepName = "MoveNext"
                Me.MoveNext()
                oPigXml.AddEleRightSign(strRow)
                If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            Loop
            Dim strValue As String = oPigXml.MainXmlStr
            Dim strRS As String = "RS"
            If RSNo > 0 Then strRS &= RSNo.ToString
            oPigXml.Clear()
            oPigXml.AddEleLeftSign(strRS, True)
            oPigXml.AddEleLeftAttribute("TotalRows", intRowNo)
            oPigXml.AddEleLeftAttribute("IsEOF", Me.EOF)
            oPigXml.AddEleLeftSignEnd()
            Dim strColInfXml As String = ""
            LOG.StepName = "mGetRSColInfXml"
            LOG.Ret = Me.mGetRSColInfXml(strColInfXml)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            oPigXml.AddXmlFragment(strColInfXml)
            oPigXml.AddEleLeftSign("Rows")
            oPigXml.AddXmlFragment(strValue)
            oPigXml.AddEleRightSign("Rows")
            oPigXml.AddEleRightSign(strRS)
            OutXml = oPigXml.MainXmlStr
            oPigXml = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mRecordset2JSon(Optional TopRows As Long = -1) As String
        Dim LOG As New PigStepLog("mRecordset2JSon")
        Try
            Dim intRowNo As Integer = 0
            LOG.StepName = "New PigJSon"
            Dim pjMain As New PigJSonLite
            If pjMain.LastErr <> "" Then Throw New Exception(pjMain.LastErr)
            pjMain.AddArrayEleBegin("ROW", True)
            Do While Not Me.EOF
                If intRowNo >= Me.MaxTopJSonOrXmlRows Then Exit Do
                If TopRows > 0 Then
                    If intRowNo >= TopRows Then Exit Do
                End If
                LOG.StepName = "Row2JSon"
                Dim strRowJSon As String = Me.Row2JSon
                If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
                If intRowNo = 0 Then
                    pjMain.AddArrayEleValue(strRowJSon, True)
                Else
                    pjMain.AddArrayEleValue(strRowJSon)
                End If
                intRowNo += 1
                LOG.StepName = "MoveNext"
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
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Convert recordset to simple JSON array, The returned result cannot be used as a standalone JSON.|当前结果集转换成简单的JSON数组，返回结果不能作为独立的 JSon 使用。
    ''' </summary>
    Public Function Recordset2SimpleJSonArray() As String
        Dim LOG As New PigStepLog("Recordset2SimpleJSonArray")
        Try
            Dim intRowNo As Integer = 0
            LOG.StepName = "New PigJSonLite"
            Dim pjMain As New PigJSonLite
            If pjMain.LastErr <> "" Then Throw New Exception(pjMain.LastErr)
            Do While Not Me.EOF
                If intRowNo >= Me.MaxTopJSonOrXmlRows Then Exit Do
                LOG.StepName = "Row2JSon"
                Dim strRowJSon As String = Me.Row2JSon
                If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
                If intRowNo = 0 Then
                    pjMain.AddArrayEleValue(strRowJSon, True)
                Else
                    pjMain.AddArrayEleValue(strRowJSon)
                End If
                intRowNo += 1
                LOG.StepName = "MoveNext"
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


    Public Function AllRecordset2Xml(ByRef OutStr As String) As String
        Dim LOG As New PigStepLog("AllRecordset2Xml")
        Try
            Dim oPigXml As New PigXml(False)
            With oPigXml
                .AddEleLeftSign("XmlRS")
                Dim strXmlRS As String = "", intRSNo As Integer = 1
                LOG.StepName = "mRecordset2Xml"
                LOG.Ret = Me.mRecordset2Xml(strXmlRS, Me.MaxTopJSonOrXmlRows, intRSNo)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(intRSNo.ToString)
                    Throw New Exception(LOG.Ret)
                End If
                .AddXmlFragment(strXmlRS)
                Dim rsParent As Recordset = Nothing
                Dim rsSub As Recordset = Me.NextRecordset
                Do While rsSub IsNot Nothing
                    intRSNo += 1
                    LOG.StepName = "mRecordset2Xml"
                    LOG.Ret = rsSub.mRecordset2Xml(strXmlRS, Me.MaxTopJSonOrXmlRows, intRSNo)
                    If LOG.Ret <> "OK" Then
                        LOG.AddStepNameInf(intRSNo.ToString)
                        Throw New Exception(LOG.Ret)
                    End If
                    .AddXmlFragment(strXmlRS)
                    rsParent = rsSub
                    LOG.StepName = "rs.NextRecordset"
                    rsSub = Nothing
                    rsSub = rsParent.NextRecordset
                    If rsParent.LastErr <> "" Then Exit Do
                Loop
                .AddEle("TotalRS", intRSNo.ToString)
                .AddEleRightSign("XmlRS")
                OutStr = .MainXmlStr
            End With
            oPigXml = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function AllRecordset2Xml(ByRef OutRS As XmlRS) As String
        Dim LOG As New PigStepLog("AllRecordset2Xml")
        Try
            Dim strXml As String = ""
            LOG.StepName = "AllRecordset2Xml(OutStr)"
            LOG.Ret = Me.AllRecordset2Xml(strXml)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "New XmlRS"
            OutRS = New XmlRS(strXml)
            If OutRS Is Nothing Then
                LOG.AddStepNameInf(strXml)
                Throw New Exception("OutRS Is Nothing")
            End If
            If OutRS.LastErr <> "" Then
                LOG.AddStepNameInf(strXml)
                Throw New Exception(OutRS.LastErr)
            End If
            Return "OK"
        Catch ex As Exception
            OutRS = Nothing
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Convert all recordset to JSON|所有结果集转换成JSON
    ''' </summary>
    ''' <returns></returns>
    Public Function AllRecordset2JSon() As String
        Dim LOG As New PigStepLog("AllRecordset2JSon")
        Try
            Dim intRSNo As Integer = 0
            LOG.StepName = "New PigJSonLite"
            Dim pjMain As New PigJSonLite
            If pjMain.LastErr <> "" Then Throw New Exception(pjMain.LastErr)
            pjMain.AddArrayEleBegin("RS", True)
            Dim strRsJSon As String
            LOG.StepName = "Me.Recordset2JSon"
            strRsJSon = Me.Recordset2JSon(Me.MaxTopJSonOrXmlRows)
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            pjMain.AddArrayEleValue(strRsJSon, True)
            intRSNo = 1
            LOG.StepName = "Me.NextRecordset"
            Dim rsParent As Recordset = Nothing
            Dim rsSub As Recordset = Me.NextRecordset
            Do While Not rsSub Is Nothing
                LOG.StepName = "rs.Recordset2JSon"
                strRsJSon = rsSub.Recordset2JSon(Me.MaxTopJSonOrXmlRows)
                If rsSub.LastErr <> "" Then Throw New Exception(rsSub.LastErr)
                pjMain.AddArrayEleValue(strRsJSon)
                intRSNo += 1
                rsParent = rsSub
                LOG.StepName = "rs.NextRecordset"
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
    Public Property IsTrimJSonValue() As Boolean
        Get
            Return Me.IsStrValueTrim
        End Get
        Set(value As Boolean)
            Me.IsStrValueTrim = value
        End Set
    End Property

    Protected Overrides Sub Finalize()
        If Not moSqlDataReader Is Nothing Then
            moSqlDataReader.Close()
        End If
        MyBase.Finalize()
    End Sub

    Private Sub mNew(SqlDataReader As SqlDataReader)
        Dim LOG As New PigStepLog("mNew")
        Try
            With Me
                LOG.StepName = "ExecuteReader"
                moSqlDataReader = SqlDataReader
                mlngRecordsAffected = moSqlDataReader.RecordsAffected
                LOG.StepName = "New Fields"
                .Fields = New Fields
                For i = 0 To moSqlDataReader.FieldCount - 1
                    LOG.StepName = "Fields.Add（" & i & ")"
                    Dim strName As String = moSqlDataReader.GetName(i)
                    Dim strTypeName As String = moSqlDataReader.GetDataTypeName(i)
                    Dim oType As Type = moSqlDataReader.GetFieldType(i)
                    Dim strFieldTypeName As String = ""
                    If oType IsNot Nothing Then strFieldTypeName = oType.Name
                    .Fields.Add(strName, strTypeName, strFieldTypeName, i)
                    If .Fields.LastErr <> "" Then Throw New Exception(.Fields.LastErr)
                Next
                If moSqlDataReader.HasRows = True Then
                    LOG.StepName = "MoveNext"
                    .MoveNext()
                    If .LastErr <> "" Then Throw New Exception(.LastErr)
                End If
            End With
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    'Public ReadOnly Property HasNextRecordset As Boolean
    '    Get
    '        Try
    '            If moSqlDataReader Is Nothing Then
    '                Return False
    '            Else
    '                Return moSqlDataReader.NextResult()
    '            End If
    '        Catch ex As Exception
    '            Me.SetSubErrInf("HasNextRecordset", ex)
    '            Return False
    '        End Try
    '    End Get
    'End Property

    Public Function NextRecordset() As Recordset
        Try
            If moSqlDataReader.NextResult() = False Then Throw New Exception("Has not next recordset")
            Dim oRecordset As Recordset = New Recordset(moSqlDataReader)
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
    ''' The maximum number of rows to convert the Recordset to JSON or XML
    ''' </summary>
    Public Property MaxTopJSonOrXmlRows As Long = 1024

    ''' <summary>
    ''' The maximum number of rows to convert the Recordset to JSON
    ''' </summary>
    Public Property MaxToJSonRows() As Long
        Get
            Return Me.MaxTopJSonOrXmlRows
        End Get
        Set(value As Long)
            Me.MaxTopJSonOrXmlRows = value
        End Set
    End Property

End Class
