'**********************************
'* Name: SQLSrvTools
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Common SQL server tools
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.31
'* Create Time: 1/9/2021
'* 1.0		1/9/2021   Add IsDBObjExists,IsDBUserExists,IsDatabaseExists,IsLoginUserExists
'* 1.1		17/9/2021   Modify IsDBObjExists,IsDBUserExists,IsDatabaseExists,IsLoginUserExists
'* 1.2		20/9/2021   Modify IsDBObjExists,IsDBUserExists,IsDatabaseExists,IsLoginUserExists
'* 1.3		5/12/2021   Add IsTabColExists
'* 1.4		6/6/2021    Imports PigToolsLiteLib
'* 1.5		9/6/2021    Add GetTableOrView2VBCode,DataCategory2VBDataType,SQLSrvTypeDataCategory
'* 1.6		13/6/2021   Modif GetTableOrView2VBCode, add DataCategory2StrValue
'* 1.7		17/6/2021   Add GetTableOrView2SQLOrVBFragment
'* 1.8		23/6/2021   Modify GetTableOrView2VBCode, add DataTypeStr,SpHelpFields2SQLSrvTypeStr
'* 1.9		25/6/2021   Modify GetTableOrView2VBCode
'* 1.10		26/6/2021   Modify GetTableOrView2VBCode
'* 1.11		1/7/2021    Modify GetTableOrView2SQLOrVBFragment
'* 1.12		2/7/2022	Use PigBaseLocal
'* 1.16		4/7/2022	Modify GetTableOrView2VBCode
'* 1.17		26/7/2022	Modify Imports
'* 1.18		28/7/2022	Modify GetTableOrView2VBCode
'* 1.19		29/7/2022	Modify Imports
'* 1.20		30/7/2022	Add mExecuteNonQuery,MkDBFunc_IsDBObjExists, modify IsDBObjExists
'* 1.21		5/8/2022	Modify GetTableOrView2VBCode
'* 1.22		16/8/2022	Modify GetTableOrView2VBCode
'* 1.23		5/9/2022	Modify datetime
'* 1.25		11/10/2022	Modify GetTableOrView2VBCode
'* 1.26		28/1/2023	Modify SpHelpFields2SQLSrvTypeStr
'* 1.27		6/3/2023	Add DropTable,ChkObjNameIsInvalid,IsReservedKeywords
'* 1.28		31/3/2023	Modify GetTableOrView2VBCode
'* 1.29		3/4/2023	Modify GetTableOrView2SQLOrVBFragment
'* 1.30		12/4/2023	Modify GetTableOrView2VBCode,GetTableOrView2SQLOrVBFragment
'* 1.31  21/7/2024  Modify PigFunc to PigFuncLite
'**********************************
Imports System.Data
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
#Else
Imports Microsoft.Data.SqlClient
#End If
Imports PigToolsLiteLib

''' <summary>
''' Common SQL Server toolsets|常用的SQL Server工具集
''' </summary>
Public Class SQLSrvTools
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1." & "31" & "." & "2"
    Private Property mConnSQLSrv As ConnSQLSrv
    Private ReadOnly Property mPigFunc As New PigFuncLite

    Public Enum EnmDBObjType
        Unknow = 0
        UserTable = 10
        View = 20
        StoredProcedure = 30
        ScalarFunction = 40
        InlineFunction = 50
        PrimaryKey = 60
        ForeignKey = 70
        Trigger = 80
        DefaultConstraint = 90
        CheckConstraint = 100
        Rule = 110
    End Enum

    Public Sub New(ConnSQLSrv As ConnSQLSrv)
        MyBase.New(CLS_VERSION)
        Try
            mConnSQLSrv = ConnSQLSrv
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    Public Function MkDBFunc_IsDBObjExists() As String
        Dim LOG As New PigStepLog("MkDBFunc_IsDBObjExists")
        Try
            Dim strFuncName As String = "_pfIsDBObjExists"
            LOG.StepName = "IsDBObjExists"
            Dim bolIsExists As Boolean = Me.IsDBObjExists(EnmDBObjType.ScalarFunction, strFuncName)
            Dim strSQL As String = ""
            If bolIsExists = True Then
                strSQL = "ALTER "
            Else
                strSQL = "CREATE "
            End If
            Me.mPigFunc.AddMultiLineText(strSQL, " Function dbo." & strFuncName & "(@DBObjType varchar(10),@ObjName sysname,@ParentObjName sysname = NULL)")
            Me.mPigFunc.AddMultiLineText(strSQL, "RETURNS bit")
            Me.mPigFunc.AddMultiLineText(strSQL, "As")
            Me.mPigFunc.AddMultiLineText(strSQL, "BEGIN")
            Me.mPigFunc.AddMultiLineText(strSQL, "Declare	@Ret bit", 1)
            Me.mPigFunc.AddMultiLineText(strSQL, "Declare @ParentObjID int = 0", 1)
            Me.mPigFunc.AddMultiLineText(strSQL, "If @ParentObjName Is Not NULL Set @ParentObjID=OBJECT_ID(@ParentObjName)", 1)
            Me.mPigFunc.AddMultiLineText(strSQL, "If @ParentObjID Is NULL", 1)
            Me.mPigFunc.AddMultiLineText(strSQL, "Set @Ret = 0", 2)
            Me.mPigFunc.AddMultiLineText(strSQL, "Else If EXISTS(Select 1 from sysobjects With(NOLOCK) where name=@ObjName And xtype=@DBObjType And parent_obj=@ParentObjID)", 1)
            Me.mPigFunc.AddMultiLineText(strSQL, "Set @Ret = 1", 2)
            Me.mPigFunc.AddMultiLineText(strSQL, "Else", 1)
            Me.mPigFunc.AddMultiLineText(strSQL, "Set @Ret = 0", 2)
            Me.mPigFunc.AddMultiLineText(strSQL, "Return(@Ret)", 1)
            Me.mPigFunc.AddMultiLineText(strSQL, "End")
            LOG.StepName = "mExecuteNonQuery"
            LOG.Ret = Me.mExecuteNonQuery(strSQL)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function MkDBFunc_IsTabColExists() As String
        Dim LOG As New PigStepLog("MkDBFunc_IsTabColExists")
        Try
            Dim strFuncName As String = "_pfIsTabColExists"
            LOG.StepName = "IsDBObjExists"
            Dim bolIsExists As Boolean = Me.IsDBObjExists(EnmDBObjType.ScalarFunction, strFuncName)
            Dim strSQL As String = ""
            If bolIsExists = True Then
                strSQL = "ALTER "
            Else
                strSQL = "CREATE "
            End If
            Me.mPigFunc.AddMultiLineText(strSQL, " Function dbo." & strFuncName & "(@TableName sysname,@ColName sysname)")
            Me.mPigFunc.AddMultiLineText(strSQL, " RETURNS bit")
            Me.mPigFunc.AddMultiLineText(strSQL, " As")
            Me.mPigFunc.AddMultiLineText(strSQL, " BEGIN")
            Me.mPigFunc.AddMultiLineText(strSQL, " Declare	@Ret bit", 1)
            Me.mPigFunc.AddMultiLineText(strSQL, " If EXISTS(Select TOP 1 1 FROM syscolumns c With(NOLOCK)  JOIN sysobjects o  With(NOLOCK) On c.id=o.id And o.xtype='U' WHERE o.name=@TableName AND c.name=@ColName)", 1)
            Me.mPigFunc.AddMultiLineText(strSQL, " SET @Ret = 1", 2)
            Me.mPigFunc.AddMultiLineText(strSQL, " ELSE", 1)
            Me.mPigFunc.AddMultiLineText(strSQL, " SET @Ret = 0", 2)
            Me.mPigFunc.AddMultiLineText(strSQL, " RETURN(@Ret)", 1)
            Me.mPigFunc.AddMultiLineText(strSQL, " END")
            LOG.StepName = "mExecuteNonQuery"
            LOG.Ret = Me.mExecuteNonQuery(strSQL)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function mExecuteNonQuery(SQL As String) As String
        Dim LOG As New PigStepLog("mExecuteNonQuery")
        Try
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(SQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                LOG.StepName = "ExecuteNonQuery"
                LOG.Ret = .ExecuteNonQuery()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End With
            oCmdSQLSrvText = Nothing
            Return "OK"
        Catch ex As Exception
            If Me.IsDebug Then LOG.AddStepNameInf(SQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Specify whether the database object exists|指定数据库对象是否存在
    ''' </summary>
    ''' <param name="DBObjType">Database Object Type|数据库对象类型</param>
    ''' <param name="ObjName">Object Name|对象名称</param>
    ''' <param name="ParentObjName">Parent Object Name|父对象名称</param>
    ''' <returns></returns>
    Public Function IsDBObjExists(DBObjType As EnmDBObjType, ObjName As String, Optional ParentObjName As String = "") As Boolean
        Const SUB_NAME As String = "IsDBObjExists"
        Dim strStepName As String = ""
        Try
            Dim strXType As String = ""
            strStepName = "Check DBObjType"
            Select Case DBObjType
                Case EnmDBObjType.UserTable
                    strXType = "U"
                Case EnmDBObjType.View
                    strXType = "V"
                Case EnmDBObjType.StoredProcedure
                    strXType = "P"
                Case EnmDBObjType.ScalarFunction
                    strXType = "FN"
                Case EnmDBObjType.InlineFunction
                    strXType = "IF"
                Case EnmDBObjType.PrimaryKey
                    strXType = "PK"
                Case EnmDBObjType.ForeignKey
                    strXType = "F"
                Case EnmDBObjType.DefaultConstraint
                    strXType = "D"
                Case EnmDBObjType.Trigger
                    strXType = "TR"
                Case EnmDBObjType.CheckConstraint
                    strXType = "C"
                Case EnmDBObjType.Rule
                    strXType = "R"
                Case Else
                    Throw New Exception("Cannot support")
            End Select
            Dim strSQL As String = "SELECT 1 FROM sysobjects WITH(NOLOCK) WHERE name=@ObjName AND xtype=@DBObjType"
            If ParentObjName <> "" Then strSQL &= " AND parent_obj=OBJECT_ID(@ParentObjName)"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                .AddPara("@ObjName", SqlDbType.VarChar, 512)
                .ParaValue("@ObjName") = ObjName
                .AddPara("@DBObjType", SqlDbType.VarChar, 10)
                .ParaValue("@DBObjType") = strXType
                If ParentObjName <> "" Then
                    .AddPara("@ParentObjName", SqlDbType.VarChar, 512)
                    .ParaValue("@ParentObjName") = ParentObjName
                End If
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
                .ActiveConnection = Me.mConnSQLSrv.Connection
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

    ''' <summary>
    ''' Whether the login name exists|登录名是否存在
    ''' </summary>
    ''' <param name="LoginName">Login Name|登录名</param>
    ''' <returns></returns>
    Public Function IsLoginUserExists(LoginName As String) As Boolean
        Const SUB_NAME As String = "IsLoginUserExists"
        Dim strStepName As String = ""
        Try
            Dim strSQL As String = "select 1 from master.dbo.syslogins WITH(NOLOCK) where name=@LoginName"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
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

    ''' <summary>
    ''' Whether the database user exists|数据库用户是否存在
    ''' </summary>
    ''' <param name="DBUserName">Database user|数据库用户</param>
    ''' <returns></returns>
    Public Function IsDBUserExists(DBUserName As String) As Boolean
        Const SUB_NAME As String = "IsDBUserExists"
        Dim strStepName As String = ""
        Try
            Dim strSQL As String = "select 1 from sysusers WITH(NOLOCK) where name=@DBUserName and islogin=1"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
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

    ''' <summary>
    ''' Whether the database table column exists|数据库表列是否存在
    ''' </summary>
    ''' <param name="TableName">TableName|表名</param>
    ''' <param name="ColName">列名|Column name</param>
    ''' <returns></returns>
    Public Function IsTabColExists(TableName As String, ColName As String) As Boolean
        Const SUB_NAME As String = "IsTabColExists"
        Dim strStepName As String = ""
        Try
            Dim strXType As String = ""
            Dim strSQL As String = "SELECT TOP 1 1 FROM syscolumns c WITH(NOLOCK)  JOIN sysobjects o  WITH(NOLOCK) ON c.id=o.id AND o.xtype='U' WHERE o.name=@TableName AND c.name=@ColName"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
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
    ''' <param name="IsSimpleProperty">是否使用简单属性代码|Whether to use simple attribute code</param>
    ''' <param name="IsSetUpdateTime">是否设置更新时间|Is set update time</param>
    ''' <returns></returns>
    Public Function GetTableOrView2VBCode(TableOrViewName As String, ByRef OutVBCode As String, Optional NotMathFillByRsList As String = "", Optional NotMathMD5List As String = "", Optional IsSimpleProperty As Boolean = True, Optional IsSetUpdateTime As Boolean = False) As String
        Dim LOG As New PigStepLog("GetTableOrView2VBCode")
        Try
            OutVBCode = "Imports PigToolsLiteLib" & Me.OsCrLf
#If NETFRAMEWORK Then
            OutVBCode &= "Imports PigSQLSrvLib" & Me.OsCrLf
#Else
            OutVBCode &= "Imports PigSQLSrvCoreLib" & Me.OsCrLf
#End If
            OutVBCode &= "Public Class " & TableOrViewName & Me.OsCrLf
            OutVBCode &= vbTab & "Inherits PigBaseLocal" & Me.OsCrLf
            OutVBCode &= vbTab & "Private Const CLS_VERSION As String = ""1.0.0""" & Me.OsCrLf

            Dim strPublic As String = ""
            Dim strProperty As String = ""
            Dim strValueMD5 As String = ""
            Dim strFillByRs As String = ""
            Dim strFillByXmlRs As String = ""
            If NotMathFillByRsList <> "" Then
                If Left(NotMathFillByRsList, 1) <> "," Then NotMathFillByRsList = "," & NotMathFillByRsList
                If Right(NotMathFillByRsList, 1) <> "," Then NotMathFillByRsList &= ","
            End If
            If NotMathMD5List <> "" Then
                If Left(NotMathMD5List, 1) <> "," Then NotMathMD5List = "," & NotMathMD5List
                If Right(NotMathMD5List, 1) <> "," Then NotMathMD5List &= ","
            End If
            LOG.StepName = "New CmdSQLSrvSp"
            Dim oCmdSQLSrvSp As New CmdSQLSrvSp("sp_help")
            With oCmdSQLSrvSp
                LOG.StepName = "Set ActiveConnection"
                .ActiveConnection = Me.mConnSQLSrv.Connection
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
                    Dim intDataCategory As Field.EnumDataCategory = Me.SQLSrvType2DataCategory(strType)
                    Dim strVBDataType As String = Me.DataCategory2VBDataType(intDataCategory)
                    Dim strValueType As String = Me.DataCategory2ValueType(intDataCategory)
                    If bolIsFrist = True Then
                        OutVBCode &= vbTab & "Public Sub New(" & strColumn_name & " As " & strVBDataType & ")" & Me.OsCrLf
                        OutVBCode &= vbTab & vbTab & "MyBase.New(CLS_VERSION)" & Me.OsCrLf
                        OutVBCode &= vbTab & vbTab & "Me." & strColumn_name & " = " & strColumn_name & Me.OsCrLf
                        OutVBCode &= vbTab & "End Sub" & Me.OsCrLf
                        strPublic &= vbTab & "Public ReadOnly Property " & strColumn_name & " As " & strVBDataType & Me.OsCrLf
                        strProperty &= vbTab & "Public ReadOnly Property " & strColumn_name & " As " & strVBDataType & Me.OsCrLf
                        If IsSimpleProperty = False And IsSetUpdateTime = True Then
                            strProperty &= vbTab & "Private mUpdateCheck As New UpdateCheck" & Me.OsCrLf
                            strProperty &= vbTab & "Public ReadOnly Property LastUpdateTime() As Date" & Me.OsCrLf
                            strProperty &= vbTab & vbTab & "Get" & Me.OsCrLf
                            strProperty &= vbTab & vbTab & vbTab & "Return mUpdateCheck.LastUpdateTime" & Me.OsCrLf
                            strProperty &= vbTab & vbTab & "End Get" & Me.OsCrLf
                            strProperty &= vbTab & "End Property" & Me.OsCrLf
                            strProperty &= vbTab & "Public ReadOnly Property IsUpdate(PropertyName As String) As Boolean" & Me.OsCrLf
                            strProperty &= vbTab & vbTab & "Get" & Me.OsCrLf
                            strProperty &= vbTab & vbTab & vbTab & "Return mUpdateCheck.IsUpdated(PropertyName)" & Me.OsCrLf
                            strProperty &= vbTab & vbTab & "End Get" & Me.OsCrLf
                            strProperty &= vbTab & "End Property" & Me.OsCrLf
                            strProperty &= vbTab & "Public ReadOnly Property HasUpdated() As Boolean" & Me.OsCrLf
                            strProperty &= vbTab & vbTab & "Get" & Me.OsCrLf
                            strProperty &= vbTab & vbTab & vbTab & "Return mUpdateCheck.HasUpdated" & Me.OsCrLf
                            strProperty &= vbTab & vbTab & "End Get" & Me.OsCrLf
                            strProperty &= vbTab & "End Property" & Me.OsCrLf
                            strProperty &= vbTab & "Public Sub UpdateCheckClear()" & Me.OsCrLf
                            strProperty &= vbTab & vbTab & "mUpdateCheck.Clear()" & Me.OsCrLf
                            strProperty &= vbTab & "End Sub" & Me.OsCrLf
                        End If
                        '-------
                        strFillByRs &= vbTab & "Friend Function fFillByRs(ByRef InRs As Recordset, Optional ByRef UpdateCnt As Integer = 0) As String" & Me.OsCrLf
                        strFillByRs &= vbTab & vbTab & "Try" & Me.OsCrLf
                        strFillByRs &= vbTab & vbTab & vbTab & "If InRs.EOF = False Then" & Me.OsCrLf
                        strFillByRs &= vbTab & vbTab & vbTab & vbTab & "With InRs.Fields" & Me.OsCrLf
                        '-------
                        strFillByXmlRs &= vbTab & "Friend Function fFillByXmlRs(ByRef InXmlRs As XmlRS, RSNo As Integer, RowNo As Integer, Optional ByRef UpdateCnt As Integer = 0) As String" & Me.OsCrLf
                        strFillByXmlRs &= vbTab & vbTab & "Try" & Me.OsCrLf
                        'strFillByXmlRs &= vbTab & vbTab & vbTab & "If InXmlRs.IsEOF(RSNo) = False Then" & Me.OsCrLf
                        strFillByXmlRs &= vbTab & vbTab & vbTab & "If RowNo <= InXmlRs.TotalRows(RSNo) Then" & Me.OsCrLf
                        strFillByXmlRs &= vbTab & vbTab & vbTab & vbTab & "With InXmlRs" & Me.OsCrLf
                        '-------
                        strValueMD5 &= vbTab & "Friend ReadOnly Property ValueMD5(Optional TextType As PigMD5.enmTextType = PigMD5.enmTextType.UTF8) As String" & Me.OsCrLf
                        strValueMD5 &= vbTab & vbTab & "Get" & Me.OsCrLf
                        strValueMD5 &= vbTab & vbTab & vbTab & "Try" & Me.OsCrLf
                        strValueMD5 &= vbTab & vbTab & vbTab & vbTab & "Dim strText As String = """"" & Me.OsCrLf
                        strValueMD5 &= vbTab & vbTab & vbTab & vbTab & "With Me" & Me.OsCrLf
                        bolIsFrist = False
                    Else
                        If IsSimpleProperty = True Then
                            strPublic &= vbTab & "Public Property " & strColumn_name & " As " & strVBDataType & Me.OsCrLf
                        Else
                            If strColumn_name <> "LastUpdateTime" Then
                                strProperty &= vbTab & "Private m" & strColumn_name & " As " & strVBDataType
                                Select Case strVBDataType
                                    Case "DateTime"
                                        strProperty &= " = #1/1/1753#"
                                    Case "String"
                                        strProperty &= " = """""
                                End Select
                                strProperty &= Me.OsCrLf
                                strProperty &= vbTab & "Public Property " & strColumn_name & "() As " & strVBDataType & Me.OsCrLf
                                strProperty &= vbTab & vbTab & "Get" & Me.OsCrLf
                                strProperty &= vbTab & vbTab & vbTab & "Return m" & strColumn_name & Me.OsCrLf
                                strProperty &= vbTab & vbTab & "End Get" & Me.OsCrLf
                                strProperty &= vbTab & vbTab & "Friend Set(value As " & strVBDataType & ")" & Me.OsCrLf
                                If IsSetUpdateTime = True Then
                                    strProperty &= vbTab & vbTab & vbTab & "If value <> m" & strColumn_name & " Then" & Me.OsCrLf
                                    strProperty &= vbTab & vbTab & vbTab & vbTab & "Me.mUpdateCheck.Add(""" & strColumn_name & """)" & Me.OsCrLf
                                    strProperty &= vbTab & vbTab & vbTab & vbTab & "m" & strColumn_name & " = value" & Me.OsCrLf
                                    strProperty &= vbTab & vbTab & vbTab & "End If" & Me.OsCrLf
                                Else
                                    strProperty &= vbTab & vbTab & vbTab & vbTab & "m" & strColumn_name & " = value" & Me.OsCrLf
                                End If
                                strProperty &= vbTab & vbTab & "End Set" & Me.OsCrLf
                                strProperty &= vbTab & "End Property" & Me.OsCrLf
                            End If
                        End If
                        If InStr(NotMathFillByRsList, "," & strColumn_name & ",") = 0 Then
                            strFillByRs &= vbTab & vbTab & vbTab & vbTab & vbTab & "If .IsItemExists(""" & strColumn_name & """) = True Then " & Me.OsCrLf
                            strFillByRs &= vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "If Me." & strColumn_name & " <> .Item(""" & strColumn_name & """)." & strValueType & " Then" & Me.OsCrLf
                            strFillByRs &= vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Me." & strColumn_name & " = .Item(""" & strColumn_name & """)." & strValueType & Me.OsCrLf
                            strFillByRs &= vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "UpdateCnt += 1" & Me.OsCrLf
                            strFillByRs &= vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "End If" & Me.OsCrLf
                            strFillByRs &= vbTab & vbTab & vbTab & vbTab & vbTab & "End If" & Me.OsCrLf
                            '--------
                            strFillByXmlRs &= vbTab & vbTab & vbTab & vbTab & vbTab & "If .IsColExists(RSNo, """ & strColumn_name & """) = True Then " & Me.OsCrLf
                            strFillByXmlRs &= vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "If Me." & strColumn_name & " <> ." & strValueType & "(RSNo, RowNo, """ & strColumn_name & """)" & " Then" & Me.OsCrLf
                            strFillByXmlRs &= vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Me." & strColumn_name & " = ." & strValueType & "(RSNo, RowNo, """ & strColumn_name & """)" & Me.OsCrLf
                            strFillByXmlRs &= vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "UpdateCnt += 1" & Me.OsCrLf
                            strFillByXmlRs &= vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "End If" & Me.OsCrLf
                            strFillByXmlRs &= vbTab & vbTab & vbTab & vbTab & vbTab & "End If" & Me.OsCrLf
                        End If
                        If InStr(NotMathMD5List, "," & strColumn_name & ",") = 0 Then
                            strValueMD5 &= vbTab & vbTab & vbTab & vbTab & vbTab & Me.GetValueMD5Row(strColumn_name, intDataCategory) & Me.OsCrLf
                        End If
                    End If
                    LOG.StepName = "MoveNext"
                    rs.MoveNext()
                    If rs.LastErr <> "" Then Throw New Exception(rs.LastErr)
                Loop
                strFillByRs &= vbTab & vbTab & vbTab & vbTab & vbTab & "Me.mUpdateCheck.Clear()" & Me.OsCrLf
                strFillByRs &= vbTab & vbTab & vbTab & vbTab & "End With" & Me.OsCrLf
                strFillByRs &= vbTab & vbTab & vbTab & "End If" & Me.OsCrLf
                strFillByRs &= vbTab & vbTab & vbTab & "Return ""OK""" & Me.OsCrLf
                strFillByRs &= vbTab & vbTab & "Catch ex As Exception" & Me.OsCrLf
                strFillByRs &= vbTab & vbTab & vbTab & "Return Me.GetSubErrInf(""fFillByRs"", ex)" & Me.OsCrLf
                strFillByRs &= vbTab & vbTab & "End Try" & Me.OsCrLf
                strFillByRs &= vbTab & "End Function" & Me.OsCrLf
                '-------
                strFillByXmlRs &= vbTab & vbTab & vbTab & vbTab & vbTab & "Me.mUpdateCheck.Clear()" & Me.OsCrLf
                strFillByXmlRs &= vbTab & vbTab & vbTab & vbTab & "End With" & Me.OsCrLf
                strFillByXmlRs &= vbTab & vbTab & vbTab & "End If" & Me.OsCrLf
                strFillByXmlRs &= vbTab & vbTab & vbTab & "Return ""OK""" & Me.OsCrLf
                strFillByXmlRs &= vbTab & vbTab & "Catch ex As Exception" & Me.OsCrLf
                strFillByXmlRs &= vbTab & vbTab & vbTab & "Return Me.GetSubErrInf(""fFillByXmlRs"", ex)" & Me.OsCrLf
                strFillByXmlRs &= vbTab & vbTab & "End Try" & Me.OsCrLf
                strFillByXmlRs &= vbTab & "End Function" & Me.OsCrLf
                '-------
                strValueMD5 &= vbTab & vbTab & vbTab & vbTab & "End With" & Me.OsCrLf
                strValueMD5 &= vbTab & vbTab & vbTab & vbTab & "Dim oPigMD5 As New PigMD5(strText, TextType)" & Me.OsCrLf
                strValueMD5 &= vbTab & vbTab & vbTab & vbTab & "ValueMD5 = oPigMD5.MD5" & Me.OsCrLf
                strValueMD5 &= vbTab & vbTab & vbTab & vbTab & "oPigMD5 = Nothing" & Me.OsCrLf
                strValueMD5 &= vbTab & vbTab & vbTab & "Catch ex As Exception" & Me.OsCrLf
                strValueMD5 &= vbTab & vbTab & vbTab & vbTab & "Me.SetSubErrInf(""ValueMD5"", ex)" & Me.OsCrLf
                strValueMD5 &= vbTab & vbTab & vbTab & vbTab & "Return """"" & Me.OsCrLf
                strValueMD5 &= vbTab & vbTab & vbTab & "End Try" & Me.OsCrLf
                strValueMD5 &= vbTab & vbTab & "End Get" & Me.OsCrLf
                strValueMD5 &= vbTab & "End Property" & Me.OsCrLf
            End With
            If IsSimpleProperty = True Then
                OutVBCode &= Me.OsCrLf & strPublic & Me.OsCrLf
            Else
                OutVBCode &= Me.OsCrLf & strProperty & Me.OsCrLf
            End If
            OutVBCode &= Me.OsCrLf & strFillByRs & Me.OsCrLf
            OutVBCode &= Me.OsCrLf & strFillByXmlRs & Me.OsCrLf
            OutVBCode &= Me.OsCrLf & strValueMD5 & Me.OsCrLf
            OutVBCode &= "End Class" & Me.OsCrLf
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function DataCategory2ValueType(DataCategory As Field.EnumDataCategory) As String
        Try
            Select Case DataCategory
                Case Field.EnumDataCategory.BooleanValue
                    DataCategory2ValueType = "BooleanValue"
                Case Field.EnumDataCategory.DateValue
                    DataCategory2ValueType = "DateValue"
                Case Field.EnumDataCategory.DecValue
                    DataCategory2ValueType = "DecValue"
                Case Field.EnumDataCategory.IntValue
                    DataCategory2ValueType = "IntValue"
                Case Field.EnumDataCategory.LongValue
                    DataCategory2ValueType = "LongValue"
                Case Field.EnumDataCategory.OtherValue
                    DataCategory2ValueType = ""
                Case Field.EnumDataCategory.StrValue
                    DataCategory2ValueType = "StrValue"
                Case Else
                    DataCategory2ValueType = ""
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("DataCategory2ValueType", ex)
            Return ""
        End Try
    End Function

    Public Function SpHelpFields2SQLSrvTypeStr(SpHelpFields As Fields) As String
        Try
            Dim strTypeName As String = SpHelpFields.Item("Type").StrValue
            SpHelpFields2SQLSrvTypeStr = strTypeName
            If SpHelpFields.Item("TrimTrailingBlanks").StrValue <> "(n/a)" Then
                Select Case strTypeName
                    Case "varchar", "char"
                    Case Else
                        SpHelpFields2SQLSrvTypeStr &= " "
                End Select
                Dim strLength As String = SpHelpFields.Item("Length").StrValue
                If strLength = "-1" Then strLength = "max"
                SpHelpFields2SQLSrvTypeStr &= "(" & strLength & ")"
            End If
        Catch ex As Exception
            Me.SetSubErrInf("SpHelpFields2SQLSrvTypeStr", ex)
            Return ""
        End Try
    End Function

    Public Function SQLSrvType2SqlDbType(SQLSrvType As String) As String
        Try
            SQLSrvType2SqlDbType = "Data.SqlDbType."
            Select Case SQLSrvType
                Case "bigint"
                    SQLSrvType2SqlDbType &= "BigInt"
                Case "binary"
                    SQLSrvType2SqlDbType &= "Binary"
                Case "bit"
                    SQLSrvType2SqlDbType &= "Bit"
                Case "char"
                    SQLSrvType2SqlDbType &= "Char"
                Case "datetime"
                    SQLSrvType2SqlDbType &= "DateTime"
                Case "decimal"
                    SQLSrvType2SqlDbType &= "Decimal"
                Case "float", "numeric"
                    SQLSrvType2SqlDbType &= "Float"
                Case "image"
                    SQLSrvType2SqlDbType &= "Image"
                Case "int"
                    SQLSrvType2SqlDbType &= "Int"
                Case "money"
                    SQLSrvType2SqlDbType &= "Money"
                Case "nchar"
                    SQLSrvType2SqlDbType &= "NChar"
                Case "ntext"
                    SQLSrvType2SqlDbType &= "NText"
                Case "nvarchar"
                    SQLSrvType2SqlDbType &= "NVarChar"
                Case "real"
                    SQLSrvType2SqlDbType &= "Real"
                Case "uniqueidentifier"
                    SQLSrvType2SqlDbType &= "UniqueIdentifier"
                Case "smalldatetime"
                    SQLSrvType2SqlDbType &= "SmallDateTime"
                Case "smallint"
                    SQLSrvType2SqlDbType &= "SmallInt"
                Case "smallmoney"
                    SQLSrvType2SqlDbType &= "SmallMoney"
                Case "text"
                    SQLSrvType2SqlDbType &= "Text"
                Case "timestamp"
                    SQLSrvType2SqlDbType &= "Timestamp"
                Case "tinyint"
                    SQLSrvType2SqlDbType &= "TinyInt"
                Case "varbinary"
                    SQLSrvType2SqlDbType &= "VarBinary"
                Case "varchar", "sysname"
                    SQLSrvType2SqlDbType &= "VarChar"
                Case "sql_variant"
                    SQLSrvType2SqlDbType &= "Variant"
                Case "xml"
                    SQLSrvType2SqlDbType &= "Xml"
                'Case ""
                '    SQLSrvType2SqlDbType &= "Udt"
                'Case ""
                '    SQLSrvType2SqlDbType &= "Structured"
                Case "date"
                    SQLSrvType2SqlDbType &= "Date"
                Case "time"
                    SQLSrvType2SqlDbType &= "Time"
                Case "datetime2"
                    SQLSrvType2SqlDbType &= "DateTime2"
                Case "datetimeoffset"
                    SQLSrvType2SqlDbType &= "DateTimeOffset"
                Case "geography", "geometry", "hierarchyid"
                    SQLSrvType2SqlDbType &= "Variant"
                Case Else
                    SQLSrvType2SqlDbType &= "Variant"
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("SQLSrvType2SqlDbType", ex)
            Return ""
        End Try
    End Function

    Public Function DataCategory2VBDataType(DataCategory As Field.EnumDataCategory) As String
        Try
            Select Case DataCategory
                Case Field.EnumDataCategory.BooleanValue
                    DataCategory2VBDataType = "Boolean"
                Case Field.EnumDataCategory.DateValue
                    DataCategory2VBDataType = "DateTime"
                Case Field.EnumDataCategory.DecValue
                    DataCategory2VBDataType = "Decimal"
                Case Field.EnumDataCategory.IntValue
                    DataCategory2VBDataType = "Integer"
                Case Field.EnumDataCategory.LongValue
                    DataCategory2VBDataType = "Long"
                Case Field.EnumDataCategory.OtherValue
                    DataCategory2VBDataType = ""
                Case Field.EnumDataCategory.StrValue
                    DataCategory2VBDataType = "String"
                Case Else
                    DataCategory2VBDataType = ""
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("DataCategory2VBDataType", ex)
            Return ""
        End Try
    End Function

    Public Function GetValueMD5Row(ColName As String, DataCategory As Field.EnumDataCategory) As String
        Try
            GetValueMD5Row = "strText &= ""<"" & "
            Select Case DataCategory
                Case Field.EnumDataCategory.BooleanValue
                    GetValueMD5Row &= "Math.Abs(CInt(." & ColName & "))"
                Case Field.EnumDataCategory.DateValue
                    GetValueMD5Row &= "Format(." & ColName & ", ""yyyy-MM-dd HH:mm:ss.fff"")"
                Case Field.EnumDataCategory.DecValue
                    GetValueMD5Row &= "Math.Round(." & ColName & ",6).ToString"
                Case Field.EnumDataCategory.OtherValue
                    GetValueMD5Row &= "." & ColName
                Case Field.EnumDataCategory.StrValue
                    GetValueMD5Row &= "." & ColName
                Case Field.EnumDataCategory.LongValue, Field.EnumDataCategory.IntValue
                    GetValueMD5Row &= "CStr(." & ColName & ")"
                Case Else
            End Select
            GetValueMD5Row &= " & "">"""
        Catch ex As Exception
            Me.SetSubErrInf("GetValueMD5Row", ex)
            Return ""
        End Try
    End Function

    Public Function SQLSrvType2DataCategory(SQLSrvType As String) As Field.EnumDataCategory
        Try
            Select Case SQLSrvType
                Case "char", "xml", "varchar", "text", "sysname", "nvarchar", "ntext", "nchar"
                    SQLSrvType2DataCategory = Field.EnumDataCategory.StrValue
                Case "tinyint", "smallint", "int"
                    SQLSrvType2DataCategory = Field.EnumDataCategory.IntValue
                Case "bigint"
                    SQLSrvType2DataCategory = Field.EnumDataCategory.LongValue
                Case "decimal", "float", "smallmoney", "real", "numeric", "money"
                    SQLSrvType2DataCategory = Field.EnumDataCategory.DecValue
                Case "date", "datetime", "datetime2", "time", "smalldatetime"
                    SQLSrvType2DataCategory = Field.EnumDataCategory.DateValue
                Case "bit"
                    SQLSrvType2DataCategory = Field.EnumDataCategory.BooleanValue
                Case "binary", "datetimeoffset", "geography", "geometry", "hierarchyid", "varbinary", "uniqueidentifier", "timestamp", "sql_variant", "image"
                    SQLSrvType2DataCategory = Field.EnumDataCategory.OtherValue
                Case Else
                    SQLSrvType2DataCategory = Field.EnumDataCategory.OtherValue
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("SQLSrvTypeDataCategory.Get", ex)
            Return Field.EnumDataCategory.OtherValue
        End Try
    End Function

    ''' <summary>
    ''' 什么SQL或VB片段|What SQL or VB Fragment
    ''' </summary>
    Public Enum EnmWhatFragment
        Unknow = 0
        ''' <summary>
        ''' 存储过程的输入参数|Input parameters of stored procedure
        ''' </summary>
        SpInParas = 1
        ''' <summary>
        ''' 存储过程的输入参数预设空值|Preset null value for input parameter of stored procedure
        ''' </summary>
        SpInParasSetNull = 2
        ''' <summary>
        ''' 调用 CmdSQLSrvSp 或 CmdSQLSrvText 的 AddPara 方法的VB代码|VB code calling AddPara of CmdSQLSrvSp or CmdSQLSrvText
        ''' </summary>
        CmdSQLSrvSpOrCmdSQLSrvText_AddPara = 3
        ''' <summary>
        ''' 调用 CmdSQLSrvSp 或 CmdSQLSrvText 的 ParaValue 方法的VB代码|VB code calling ParaValue of CmdSQLSrvSp or CmdSQLSrvText
        ''' </summary>
        CmdSQLSrvSpOrCmdSQLSrvText_ParaValue = 4
        ''' <summary>
        ''' 调用 CmdSQLSrvSp 或 CmdSQLSrvText 的 AddPara 和 ParaValue 方法的VB代码|VB code calling AddPara and ParaValue of CmdSQLSrvSp or CmdSQLSrvText
        ''' </summary>
        CmdSQLSrvSpOrCmdSQLSrvText_AddPara_ParaValue = 5
        ''' <summary>
        ''' 生成调用 CmdSQLSrvSp 或 CmdSQLSrvText 的 UPDATE SQL 语句 的VB代码|VB code calling AddPara and ParaValue of CmdSQLSrvSp or CmdSQLSrvText
        ''' </summary>
        CmdSQLSrvSpOrCmdSQLSrvText_UpdCols = 8
        ''' <summary>
        ''' 每列判断并更新|The columns of each data table are judged and updated
        ''' </summary>
        UpdatePerCol = 6
        ''' <summary>
        ''' 调用存储过程的参数列表|Parameter list of calling stored procedure
        ''' </summary>
        ExecSpParas = 7
    End Enum


    ''' <summary>
    ''' 生成表或视图对应的SQL语句或VB代码片段|Generate SQL statement or VB code fragments corresponding to tables or views
    ''' </summary>
    ''' <param name="TableOrViewName">表或视图名|Table or view name</param>
    ''' <param name="WhatFragment">什么片段|What Fragment</param>
    ''' <param name="OutFragment">输出的SQL语句片段|Output SQL statement fragment</param>
    ''' <param name="NotMathColList">不需要的列名列表，以,分隔|List of unwanted column names, separated by ","</param>
    ''' <returns></returns>
    Public Function GetTableOrView2SQLOrVBFragment(TableOrViewName As String, WhatFragment As EnmWhatFragment, ByRef OutFragment As String, Optional NotMathColList As String = "") As String
        Dim LOG As New PigStepLog("GetTableOrView2SQLOrVBFragment")
        Try
            If NotMathColList <> "" Then
                If Left(NotMathColList, 1) <> "," Then NotMathColList = "," & NotMathColList
                If Right(NotMathColList, 1) <> "," Then NotMathColList &= ","
            End If
            OutFragment = ""
            Select Case WhatFragment
                Case EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_AddPara, EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_AddPara_ParaValue, EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_ParaValue
                    OutFragment &= "Dim oCmdSQLSrvSp As New CmdSQLSrvSp(""SpName"")" & Me.OsCrLf
                    OutFragment &= "Dim oCmdSQLSrvText As New CmdSQLSrvText(""TextName"")" & Me.OsCrLf
                    OutFragment &= "With oCmdSQLSrvSpOrCmdSQLSrvText" & Me.OsCrLf
                Case EnmWhatFragment.UpdatePerCol
                    OutFragment &= "SET @Rows=0" & Me.OsCrLf
                Case EnmWhatFragment.ExecSpParas
                    OutFragment &= "EXEC SpName "
                Case EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_UpdCols
                    OutFragment &= "Dim strUpdCols As String = """"" & Me.OsCrLf
            End Select
            LOG.StepName = "New CmdSQLSrvSp"
            Dim oCmdSQLSrvSp As New CmdSQLSrvSp("sp_help")
            With oCmdSQLSrvSp
                LOG.StepName = "Set ActiveConnection"
                .ActiveConnection = Me.mConnSQLSrv.Connection
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
                Do While Not rs.EOF
                    Dim strColName As String = rs.Fields.Item("Column_name").StrValue
                    Dim strLength As String = rs.Fields.Item("Length").StrValue
                    Dim strSQLSrvType As String = rs.Fields.Item("Type").StrValue
                    Dim strSqlDbType As String = Me.SQLSrvType2SqlDbType(strSQLSrvType)
                    Dim strSQLSrvTypeStr As String = Me.SpHelpFields2SQLSrvTypeStr(rs.Fields)
                    Select Case WhatFragment
                        Case EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_AddPara
                            OutFragment &= vbTab & ".AddPara(""@" & strColName & """, " & strSqlDbType & ")" & Me.OsCrLf
                        Case EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_ParaValue
                            OutFragment &= vbTab & ".ParaValue(""@" & strColName & """) = InObj." & strColName & Me.OsCrLf
                        Case EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_AddPara_ParaValue
                            OutFragment &= vbTab & "If InObj.IsUpdate(""" & strColName & """) = True Then" & Me.OsCrLf
                            OutFragment &= vbTab & vbTab & ".AddPara(""@" & strColName & """, " & strSqlDbType
                            Select Case strSQLSrvType
                                Case "varchar"
                                    OutFragment &= " , " & strLength
                            End Select
                            OutFragment &= ")" & Me.OsCrLf
                            OutFragment &= vbTab & vbTab & ".ParaValue(""@" & strColName & """) = InObj." & strColName & Me.OsCrLf
                            OutFragment &= vbTab & "End If" & Me.OsCrLf
                        Case EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_UpdCols
                            OutFragment &= "If InObj.IsUpdate(""" & strColName & """) = True Then strUpdCols &= ""," & strColName & "=@" & strColName & """" & Me.OsCrLf
                        Case EnmWhatFragment.SpInParas, EnmWhatFragment.SpInParasSetNull
                            OutFragment &= vbTab & ",@" & strColName & " " & strSQLSrvTypeStr
                            Select Case WhatFragment
                                Case EnmWhatFragment.SpInParasSetNull
                                    OutFragment &= " = NULL"
                            End Select
                            OutFragment &= Me.OsCrLf
                        Case EnmWhatFragment.UpdatePerCol
                            OutFragment &= "IF @" & strColName & " IS NOT NULL" & Me.OsCrLf
                            OutFragment &= "BEGIN" & Me.OsCrLf
                            OutFragment &= vbTab & "UPDATE TableName SET " & strColName & "=@" & strColName & " WHERE KeyID=@KeyID" & Me.OsCrLf
                            OutFragment &= vbTab & "SET @Rows=@Rows+@@ROWCOUNT" & Me.OsCrLf
                            OutFragment &= "END" & Me.OsCrLf
                        Case EnmWhatFragment.ExecSpParas
                            OutFragment &= "@" & strColName & " = @" & strColName & ","
                    End Select
                    LOG.StepName = "MoveNext"
                    rs.MoveNext()
                    If rs.LastErr <> "" Then Throw New Exception(rs.LastErr)
                Loop
                Select Case WhatFragment
                    Case EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_AddPara, EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_AddPara_ParaValue, EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_ParaValue
                        OutFragment &= "End With" & Me.OsCrLf
                    Case EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_UpdCols
                        OutFragment &= "If strUpdCols = """" Then Throw New Exception(""There is nothing to update"")" & Me.OsCrLf
                        OutFragment &= "strUpdCols = Mid(strUpdCols, 2)" & Me.OsCrLf
                        OutFragment &= "strSQL = ""UPDATE dbo.TabName SET "" & strUpdCols & "" WHERE KeyID=@KeyID""" & Me.OsCrLf
                End Select
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function DropTable(TabName As String) As String
        Dim LOG As New PigStepLog("DropTable")
        Dim strSQL As String = ""
        Try
            If Me.IsDBObjExists(EnmDBObjType.UserTable, TabName) = False Then Throw New Exception(TabName & " not exists.")
            strSQL = "DROP TABLE dbo." & TabName
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                LOG.StepName = "ExecuteNonQuery"
                LOG.Ret = .ExecuteNonQuery()
                If LOG.Ret <> "OK" Then
                    Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
                    Throw New Exception(LOG.Ret)
                End If
            End With
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 是否保留关键字|Keep keyword or not
    ''' </summary>
    ''' <param name="ObjName"></param>
    ''' <returns></returns>
    Public Function IsReservedKeywords(ObjName As String) As Boolean
        Const RKLIST As String = "<create><table><insert><into><alues><delete><from><update><set><where><drop><alteradd><select><from><distinct><all><and><or><not><left><right><join><outer><cross><inner><using><inner><full><on><as><order><by><desc><asc><between><union><intersect><except><is><null><distinct><having>"
        Try
            ObjName = "<" & ObjName & ">"
            If InStr(RKLIST, ObjName) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 检查是否无效的对象名|Check for invalid object names
    ''' </summary>
    ''' <param name="ObjName"></param>
    ''' <returns>Returning OK means passing</returns>
    Public Function ChkObjNameIsInvalid(ObjName As String) As String
        Const SPECIAL_CHARACTERS As String = "~!%6&*()-+`={}[];',./:""<>? "
        Try
            Select Case Len(ObjName)
                Case 1 To 128
                Case Else
                    Throw New Exception("Invalid object name length")
            End Select
            Select Case Left(ObjName, 1)
                Case "0" To "9"
                    Throw New Exception("Object name cannot start with a number")
            End Select
            For i = 0 To Len(ObjName) - 1
                If InStr(Mid(SPECIAL_CHARACTERS, ObjName, i), 1) > 0 Then Throw New Exception("The object name contains special characters.")
            Next
            If Me.IsReservedKeywords(ObjName) = True Then Throw New Exception("")
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

End Class
