'**********************************
'* Name: SQLSrvTools
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Common SQL server tools
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 1/9/2021
'* 1.0		1/9/2021   Add IsDBObjExists,IsDBUserExists,IsDatabaseExists,IsLoginUserExists
'* 1.1		17/9/2021   Modify IsDBObjExists,IsDBUserExists,IsDatabaseExists,IsLoginUserExists
'* 1.2		20/9/2021   Modify IsDBObjExists,IsDBUserExists,IsDatabaseExists,IsLoginUserExists
'* 1.3		5/12/2021   Add IsTabColExists
'* 1.4		6/6/2021    Imports PigToolsLiteLib
'* 1.5		9/6/2021   Add GetTableOrView2VBCode,DataCategory2VBDataType,SQLSrvTypeDataCategory
'**********************************
Imports System.Data
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
#Else
Imports Microsoft.Data.SqlClient
Imports PigToolsLiteLib

#End If
Public Class SQLSrvTools
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.5.30"
    Private moConnSQLSrv As ConnSQLSrv

    Public Enum enmDBObjType
        Unknow = 0
        UserTable = 10
        View = 20
        StoredProcedure = 30
        ScalarFunction = 40
        InlineFunction = 50
    End Enum

    Public Sub New(ConnSQLSrv As ConnSQLSrv)
        MyBase.New(CLS_VERSION)
        Try
            moConnSQLSrv = ConnSQLSrv
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    Public Function IsDBObjExists(DBObjType As enmDBObjType, ObjName As String) As Boolean
        Const SUB_NAME As String = "IsDBObjExists"
        Dim strStepName As String = ""
        Try
            Dim strXType As String = ""
            strStepName = "Check DBObjType"
            Select Case DBObjType
                Case enmDBObjType.UserTable
                    strXType = "U"
                Case enmDBObjType.View
                    strXType = "V"
                Case enmDBObjType.StoredProcedure
                    strXType = "P"
                Case enmDBObjType.ScalarFunction
                    strXType = "FN"
                Case enmDBObjType.InlineFunction
                    strXType = "IF"
                Case Else
                    Throw New Exception("Cannot support")
            End Select
            Dim strSQL As String = "select 1 from sysobjects WITH(NOLOCK) where name=@ObjName and xtype=@DBObjType"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@ObjName", SqlDbType.VarChar, 512)
                .AddPara("@DBObjType", SqlDbType.VarChar, 10)
                .ParaValue("@ObjName") = ObjName
                .ParaValue("@DBObjType") = strXType
                strStepName = "Execute"
                Dim rsAny = .Execute()
                If .LastErr <> "" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(.LastErr)
                End If
                If rsAny.EOF = True Then
                    IsDBObjExists = False
                Else
                    IsDBObjExists = True
                End If
                strStepName = "rsAny.Close"
                rsAny.Close()
                rsAny = Nothing
            End With
            oCmdSQLSrvText = Nothing
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return False
        End Try
    End Function


    Public Function IsDatabaseExists(DBName As String) As Boolean
        Const SUB_NAME As String = "IsDatabaseExists"
        Dim strStepName As String = ""
        Try
            Dim strSQL As String = "select 1 from master.dbo.sysdatabases WITH(NOLOCK) where name=@DBName"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@DBName", SqlDbType.VarChar, 512)
                .ParaValue("@DBName") = DBName
                strStepName = "Execute"
                Dim rsAny = .Execute()
                If .LastErr <> "" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(.LastErr)
                End If
                If rsAny.EOF = True Then
                    Return False
                Else
                    Return True
                End If
                strStepName = "rsAny.Close"
                rsAny.Close()
                rsAny = Nothing
            End With
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return False
        End Try
    End Function

    Public Function IsLoginUserExists(LoginName As String) As Boolean
        Const SUB_NAME As String = "IsLoginUserExists"
        Dim strStepName As String = ""
        Try
            Dim strSQL As String = "select 1 from master.dbo.syslogins WITH(NOLOCK) where name=@LoginName"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@LoginName", SqlDbType.VarChar, 512)
                .ParaValue("@LoginName") = LoginName
                strStepName = "Execute"
                Dim rsAny = .Execute()
                If .LastErr <> "" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(.LastErr)
                End If
                If rsAny.EOF = True Then
                    Return False
                Else
                    Return True
                End If
                strStepName = "rsAny.Close"
                rsAny.Close()
                rsAny = Nothing
            End With
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return False
        End Try
    End Function

    Public Function IsDBUserExists(DBUserName As String) As Boolean
        Const SUB_NAME As String = "IsDBUserExists"
        Dim strStepName As String = ""
        Try
            Dim strSQL As String = "select 1 from sysusers WITH(NOLOCK) where name=@DBUserName and islogin=1"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@DBUserName", SqlDbType.VarChar, 512)
                .ParaValue("@DBUserName") = DBUserName
                strStepName = "Execute"
                Dim rsAny = .Execute()
                If .LastErr <> "" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(.LastErr)
                End If
                If rsAny.EOF = True Then
                    Return False
                Else
                    Return True
                End If
                strStepName = "rsAny.Close"
                rsAny.Close()
                rsAny = Nothing
            End With
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return False
        End Try
    End Function

    Public Function IsTabColExists(TableName As String, ColName As String) As Boolean
        Const SUB_NAME As String = "IsTabColExists"
        Dim strStepName As String = ""
        Try
            Dim strXType As String = ""
            Dim strSQL As String = "SELECT TOP 1 1 FROM syscolumns c WITH(NOLOCK)  JOIN sysobjects o  WITH(NOLOCK) ON c.id=o.id AND o.xtype='U' WHERE o.name=@TableName AND c.name=@ColName"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@TableName", SqlDbType.VarChar, 512)
                .AddPara("@ColName", SqlDbType.VarChar, 512)
                .ParaValue("@TableName") = TableName
                .ParaValue("@ColName") = ColName
                strStepName = "Execute"
                Dim rsAny = .Execute()
                If .LastErr <> "" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(.LastErr)
                End If
                If rsAny.EOF = True Then
                    IsTabColExists = False
                Else
                    IsTabColExists = True
                End If
                strStepName = "rsAny.Close"
                rsAny.Close()
                rsAny = Nothing
            End With
            oCmdSQLSrvText = Nothing
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 生成表或视图对应的VB类代码|Generate VB class code corresponding to table or view.
    ''' </summary>
    ''' <param name="TableOrViewName">表或视图名|Table or view name</param>
    ''' <param name="OutVBCode">输出的VB代码|Exported VB code</param>
    ''' <param name="NotMathMD5List">不匹配MD5的列名列表，以,分隔|List of column names that do not match MD5, separated by ","</param>
    ''' <returns></returns>
    Public Function GetTableOrView2VBCode(TableOrViewName As String, ByRef OutVBCode As String, Optional NotMathFillByRsList As String = "", Optional NotMathMD5List As String = "") As String
        Dim LOG As New PigStepLog("GetTableOrView2VBCode")
        Try
            OutVBCode = "Imports PigToolsLiteLib" & vbCrLf
#If NETFRAMEWORK Then
            OutVBCode &= "Imports PigSQLSrvLib" & vbCrLf
#Else
            OutVBCode &= "Imports PigSQLSrvCoreLib" & vbCrLf
#End If
            OutVBCode &= "Public Class " & TableOrViewName & vbCrLf
            OutVBCode &= vbTab & "Inherits PigBaseMini" & vbCrLf
            OutVBCode &= vbTab & "Private Const CLS_VERSION As String = ""1.0.0""" & vbCrLf

            Dim strPublic As String = ""
            Dim strValueMD5 As String = ""
            Dim strFillByRs As String = ""
            If NotMathFillByRsList <> "" Then
                If Left(NotMathFillByRsList, 1) <> "," Then NotMathFillByRsList = "," & NotMathFillByRsList
                If Right(NotMathFillByRsList, 1) <> "," Then NotMathFillByRsList &= ","
            End If
            If NotMathFillByRsList <> "" Then
                If Left(NotMathFillByRsList, 1) <> "," Then NotMathFillByRsList = "," & NotMathFillByRsList
                If Right(NotMathFillByRsList, 1) <> "," Then NotMathFillByRsList &= ","
            End If
            LOG.StepName = "New CmdSQLSrvSp"
            Dim oCmdSQLSrvSp As New CmdSQLSrvSp("sp_help")
            With oCmdSQLSrvSp
                LOG.StepName = "Set ActiveConnection"
                .ActiveConnection = Me.moConnSQLSrv.Connection
                If .LastErr <> "" Then
                    LOG.AddStepNameInf("sp_help")
                    Throw New Exception(.LastErr)
                End If
                .AddPara("@objname", SqlDbType.NVarChar)
                .ParaValue("@objname") = TableOrViewName
                LOG.StepName = "New CmdSQLSrvSp"
                Dim rs As Recordset = .Execute()
                If .LastErr <> "" Then
                    LOG.AddStepNameInf(.DebugStr)
                    Throw New Exception(.LastErr)
                End If
                LOG.StepName = "NextRecordset"
                rs = rs.NextRecordset
                Dim bolIsFrist As Boolean = True
                Do While Not rs.EOF
                    Dim strColumn_name As String = rs.Fields.Item("Column_name").StrValue
                    Dim strType As String = rs.Fields.Item("Type").StrValue
                    Dim intDataCategory As Field.DataCategoryEnum = Me.SQLSrvTypeDataCategory(strType)
                    Dim strVBDataType As String = Me.DataCategory2VBDataType(intDataCategory)
                    Dim strValueType As String = Me.DataCategory2ValueType(intDataCategory)
                    If bolIsFrist = True Then
                        OutVBCode &= vbTab & "Public Sub New(" & strColumn_name & " As " & strVBDataType & ")" & vbCrLf
                        OutVBCode &= vbTab & vbTab & "MyBase.New(CLS_VERSION)" & vbCrLf
                        OutVBCode &= vbTab & vbTab & "Me." & strColumn_name & " = " & strColumn_name & vbCrLf
                        OutVBCode &= vbTab & "End Sub" & vbCrLf
                        strPublic &= vbTab & "Public ReadOnly Property " & strColumn_name & " As " & strVBDataType & vbCrLf
                        strFillByRs &= vbTab & "Friend Function fFillByRs(ByRef InRs As Recordset) As String" & vbCrLf
                        strFillByRs &= vbTab & vbTab & "Try" & vbCrLf
                        strFillByRs &= vbTab & vbTab & vbTab & "If InRs.EOF = True Then Throw New Exception(""No data"")" & vbCrLf
                        strFillByRs &= vbTab & vbTab & vbTab & "With InRs.Fields" & vbCrLf
                        bolIsFrist = False
                    Else
                        strPublic &= vbTab & "Public Property " & strColumn_name & " As " & strVBDataType & vbCrLf
                        If InStr(NotMathFillByRsList, "," & strColumn_name & ",") = 0 Then
                            strFillByRs &= vbTab & vbTab & vbTab & vbTab & "If .IsItemExists(""" & strColumn_name & """) = True Then Me." & strColumn_name & " = .Item(""" & strColumn_name & """)." & strValueType & vbCrLf
                        End If
                    End If
                    LOG.StepName = "MoveNext"
                    rs.MoveNext()
                    If rs.LastErr <> "" Then Throw New Exception(rs.LastErr)
                Loop
                strFillByRs &= vbTab & vbTab & vbTab & "End With" & vbCrLf
                strFillByRs &= vbTab & vbTab & vbTab & "Return ""OK""" & vbCrLf
                strFillByRs &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
                strFillByRs &= vbTab & vbTab & vbTab & "Return Me.GetSubErrInf(""fFillByRs"", ex)" & vbCrLf
                strFillByRs &= vbTab & vbTab & "End Try" & vbCrLf
                strFillByRs &= vbTab & "End Function" & vbCrLf
            End With
            OutVBCode &= vbCrLf & strPublic & vbCrLf
            OutVBCode &= vbCrLf & strFillByRs & vbCrLf
            OutVBCode &= "End Class" & vbCrLf
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function DataCategory2ValueType(DataCategory As Field.DataCategoryEnum) As String
        Try
            Select Case DataCategory
                Case Field.DataCategoryEnum.BooleanValue
                    DataCategory2ValueType = "BooleanValue"
                Case Field.DataCategoryEnum.DateValue
                    DataCategory2ValueType = "DateValue"
                Case Field.DataCategoryEnum.DecValue
                    DataCategory2ValueType = "DecValue"
                Case Field.DataCategoryEnum.IntValue
                    DataCategory2ValueType = "IntValue"
                Case Field.DataCategoryEnum.LongValue
                    DataCategory2ValueType = "LngValue"
                Case Field.DataCategoryEnum.OtherValue
                    DataCategory2ValueType = ""
                Case Field.DataCategoryEnum.StrValue
                    DataCategory2ValueType = "StrValue"
                Case Else
                    DataCategory2ValueType = ""
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("DataCategory2ValueType", ex)
            Return ""
        End Try
    End Function

    Public Function DataCategory2VBDataType(DataCategory As Field.DataCategoryEnum) As String
        Try
            Select Case DataCategory
                Case Field.DataCategoryEnum.BooleanValue
                    DataCategory2VBDataType = "Boolean"
                Case Field.DataCategoryEnum.DateValue
                    DataCategory2VBDataType = "DateTime"
                Case Field.DataCategoryEnum.DecValue
                    DataCategory2VBDataType = "Decimal"
                Case Field.DataCategoryEnum.IntValue
                    DataCategory2VBDataType = "Integer"
                Case Field.DataCategoryEnum.LongValue
                    DataCategory2VBDataType = "Long"
                Case Field.DataCategoryEnum.OtherValue
                    DataCategory2VBDataType = ""
                Case Field.DataCategoryEnum.StrValue
                    DataCategory2VBDataType = "String"
                Case Else
                    DataCategory2VBDataType = ""
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("DataCategory2VBDataType", ex)
            Return ""
        End Try
    End Function

    Public Function SQLSrvTypeDataCategory(SQLSrvType As String) As Field.DataCategoryEnum
        Try
            Select Case SQLSrvType
                Case "char", "xml", "varchar", "text", "sysname", "nvarchar", "ntext", "nchar"
                    SQLSrvTypeDataCategory = Field.DataCategoryEnum.StrValue
                Case "tinyint", "smallint", "int"
                    SQLSrvTypeDataCategory = Field.DataCategoryEnum.IntValue
                Case "bigint"
                    SQLSrvTypeDataCategory = Field.DataCategoryEnum.LongValue
                Case "decimal", "float", "smallmoney", "real", "numeric", "money"
                    SQLSrvTypeDataCategory = Field.DataCategoryEnum.DecValue
                Case "date", "datetime", "datetime2", "time", "smalldatetime"
                    SQLSrvTypeDataCategory = Field.DataCategoryEnum.DateValue
                Case "bit"
                    SQLSrvTypeDataCategory = Field.DataCategoryEnum.BooleanValue
                Case "binary", "datetimeoffset", "geography", "geometry", "hierarchyid", "varbinary", "uniqueidentifier", "timestamp", "sql_variant", "image"
                    SQLSrvTypeDataCategory = Field.DataCategoryEnum.OtherValue
                Case Else
                    SQLSrvTypeDataCategory = Field.DataCategoryEnum.OtherValue
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("SQLSrvTypeDataCategory.Get", ex)
            Return Field.DataCategoryEnum.OtherValue
        End Try
    End Function


End Class
