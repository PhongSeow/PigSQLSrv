'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: ConsoleDemo for PigSQLSrv
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.18.3
'* Create Time: 17/4/2021
'* 1.2	23/9/2021	Add Test Cache Query
'* 1.3	5/10/2021	Imports PigKeyCacheLib
'* 1.4	8/10/2021	Add Test Cache Query -> CmdSQLSrvSp
'* 1.5	9/10/2021	Add Test Cache Query -> Print 
'* 1.6	5/12/2021	Add Test Cache Query -> Print 
'* 1.7	15/12/2021	Test the new class library
'* 1.8	23/1/2022	Refer to PigConsole.Getpwdstr of PigCmdLib  is used to hide the entered password.
'* 1.9	2/2/2022	Add Database connection management
'* 1.10	19/3/2022	Use PigCmdLib.GetLine
'* 1.11	23/3/2022	Modify MainSet
'* 1.13	29/4/2022	Modify MainSet,Main
'* 1.14	30/4/2022	Modify MainSet
'* 1.15	1/5/2022	Modify MainSet
'* 1.16	9/6/2022	Add SQLSrvToolsDemo
'* 1.17	23/6/2022	Modify SQLSrvToolsDemo
'* 1.18	24/6/2022	Modify SQLSrvToolsDemo
'**********************************
Imports System.Data
Imports PigKeyCacheLib
Imports PigCmdLib
#If NETFRAMEWORK Then
Imports PigSQLSrvLib
Imports System.Data.SqlClient
#Else
Imports PigSQLSrvCoreLib
Imports Microsoft.Data.SqlClient
#End If
Imports PigToolsLiteLib


Public Class ConsoleDemo
    Public ConnSQLSrv As ConnSQLSrv
    Public CmdSQLSrvSp As CmdSQLSrvSp
    Public CmdSQLSrvText As CmdSQLSrvText
    Public ConnStr As String
    Public SQL As String
    Public RS As Recordset
    Public RS2 As Recordset
    Public DBSrv As String = "localhost"
    Public MirrDBSrv As String = "localhost"
    Public DBUser As String = "sa"
    Public DBPwd As String = ""
    Public CurrDB As String = "master"
    Public CurrConsoleKey As ConsoleKey
    Public InpStr As String
    Public AccessFilePath As String
    Public TableName As String
    Public ColName As String
    Public PigConsole As New PigConsole
    'Public DBConnMgr As DBConnMgr
    Public EncKey As String
    Public ConfFilePath As String
    Public Ret As String
    Public PigFunc As New PigFunc
    'Public DBConnDef As DBConnDef
    Public DBConnName As String
    Public MenuKey As String
    Public MenuKey2 As String
    Public MenuDefinition As String
    Public MenuDefinition2 As String
    Public SQLSrvTools As SQLSrvTools
    Public VBCodeOrSQLFragment As String
    Public NotMathFillByRsList As String
    Public NotMathMD5List As String
    Public NotMathColList As String
    Public FilePath As String
    Public WhatFragment As SQLSrvTools.EnmWhatFragment

    Public Sub MainFunc()
        Do While True
            Console.Clear()
            Console.WriteLine("*******************")
            Console.WriteLine("Main Function menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Up")
            Console.WriteLine("Press A to Set SQL Server Connection String")
            Console.WriteLine("Press B to OpenOrKeepActive Connection")
            Console.WriteLine("Press C to Show Connection Information")
            Console.WriteLine("Press D to Create Recordset with Execute")
            Console.WriteLine("Press E to Show Recordset Information")
            Console.WriteLine("Press F to Recordset.MoveNext")
            Console.WriteLine("Press G to Recordset.NextRecordset")
            Console.WriteLine("Press H to Test ExecuteNonQuery")
            Console.WriteLine("Press I to Test JSon")
            Console.WriteLine("Press J to Execute SQL Server StoredProcedure")
            Console.WriteLine("Press K to Execute SQL Server SQL statement Text")
            Console.WriteLine("Press L to SQLSrvTools")
            Console.WriteLine("Press M to Test MultipleActiveResultSets")
            Console.WriteLine("Press N to Test Cache Query")
            Console.CursorVisible = False
            Console.WriteLine("*******************")
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("Set Connection String")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Press Q to Up")
                    Console.WriteLine("Press A to SQL Server(StandAlone mode)")
                    Console.WriteLine("Press B to SQL Server(Mirror mode)")
                    Do While True
                        Me.CurrConsoleKey = Console.ReadKey(True).Key
                        Select Case Me.CurrConsoleKey
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Console.CursorVisible = True
                                Me.PigConsole.GetLine("Input SQL Server", Me.DBSrv)
                                Console.WriteLine("SQL Server=" & Me.DBSrv)
                                Console.WriteLine("Input Default DB:" & Me.CurrDB)
                                Me.CurrDB = Console.ReadLine()
                                If Me.CurrDB = "" Then Me.CurrDB = "master"
                                Console.WriteLine("Default DB=" & Me.CurrDB)
                                Console.WriteLine("Is Trusted Connection ? (Y/n)")
                                Me.InpStr = Console.ReadLine()
                                Select Case Me.InpStr
                                    Case "Y", "y", ""
                                        Me.ConnSQLSrv = New ConnSQLSrv(Me.DBSrv, Me.CurrDB)
                                    Case Else
                                        Console.WriteLine("Input DB User:" & Me.DBUser)
                                        Me.DBUser = Console.ReadLine()
                                        If Me.DBUser = "" Then Me.DBUser = "sa"
                                        Console.WriteLine("DB User=" & Me.DBUser)
                                        Console.WriteLine("Input DB Password:")
                                        Me.DBPwd = Me.PigConsole.GetPwdStr
                                        'Console.WriteLine("DB Password=" & Me.DBPwd)
                                        Me.ConnSQLSrv = New ConnSQLSrv(Me.DBSrv, Me.CurrDB, Me.DBUser, Me.DBPwd)
                                End Select
                                Me.ConnSQLSrv.ConnectionTimeout = 5
                                Exit Do
                            Case ConsoleKey.B
                                Console.WriteLine("Input Principal SQLServer:" & Me.DBSrv)
                                Me.DBSrv = Console.ReadLine()
                                If Me.DBSrv = "" Then Me.DBSrv = "localhost"
                                Console.WriteLine("Principal SQLServer=" & Me.DBSrv)
                                Console.WriteLine("Input Mirror SQLServer:" & Me.MirrDBSrv)
                                Me.MirrDBSrv = Console.ReadLine()
                                If Me.MirrDBSrv = "" Then Me.MirrDBSrv = "localhost"
                                Console.WriteLine("MirrorSQLServer SQLServer=" & Me.MirrDBSrv)
                                Console.WriteLine("Input Default DB:" & Me.CurrDB)
                                Me.CurrDB = Console.ReadLine()
                                If Me.CurrDB = "" Then Me.CurrDB = "master"
                                Console.WriteLine("Default DB=" & Me.CurrDB)
                                Console.WriteLine("Is Trusted Connection ? (Y/n)")
                                Me.InpStr = Console.ReadLine()
                                Select Case Me.InpStr
                                    Case "Y", "y", ""
                                        Me.ConnSQLSrv = New ConnSQLSrv(Me.DBSrv, Me.MirrDBSrv, Me.CurrDB)
                                    Case Else
                                        Console.WriteLine("Input DB User:" & Me.DBUser)
                                        Me.DBUser = Console.ReadLine()
                                        If Me.DBUser = "" Then Me.DBUser = "sa"
                                        Console.WriteLine("DB User=" & Me.DBUser)
                                        Console.WriteLine("Input DB Password:")
                                        Me.DBPwd = Me.PigConsole.GetPwdStr
                                        'Console.WriteLine("DB Password=" & Me.DBPwd)
                                        Me.ConnSQLSrv = New ConnSQLSrv(Me.DBSrv, Me.MirrDBSrv, Me.CurrDB, Me.DBUser, Me.DBPwd)
                                End Select
                                Exit Do
                        End Select
                    Loop
                Case ConsoleKey.B
                    Console.WriteLine("#################")
                    Console.WriteLine("OpenOrKeepActive Connection")
                    Console.WriteLine("#################")
                    With Me.ConnSQLSrv
                        Console.WriteLine("OpenOrKeepActive:")
                        .OpenOrKeepActive()
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                        End If
                    End With
                Case ConsoleKey.C
                    Console.WriteLine("#################")
                    Console.WriteLine("Show Connection Information")
                    Console.WriteLine("#################")
                    Console.WriteLine("ConnectionString=" & Me.ConnSQLSrv.Connection.ConnectionString)
                    Console.WriteLine("State=" & Me.ConnSQLSrv.Connection.State)
                    Console.WriteLine("ConnStatus=" & Me.ConnSQLSrv.ConnStatus)
                    Console.WriteLine("IsDBConnReady=" & Me.ConnSQLSrv.IsDBConnReady)
                Case ConsoleKey.D
                    Console.WriteLine("#################")
                    Console.WriteLine("Create Recordset with Execute")
                    Console.WriteLine("#################")
                    Console.CursorVisible = True
                    Console.WriteLine("Input SQL:")
                    Me.SQL = Console.ReadLine()
                    If Not Me.RS Is Nothing Then
                        Me.RS.Close()
                    End If
                    Me.CmdSQLSrvText = Nothing
                    Me.CmdSQLSrvText = New CmdSQLSrvText(Me.SQL)
                    With Me.CmdSQLSrvText
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            .ActiveConnection = Me.ConnSQLSrv.Connection
                            Me.RS = .Execute()
                            If .LastErr <> "" Then
                                Console.WriteLine(.LastErr)
                            Else
                                Console.WriteLine("OK")
                            End If
                        End If
                        Console.WriteLine("RecordsAffected=" & .RecordsAffected)
                    End With
                Case ConsoleKey.E
                    Console.WriteLine("#################")
                    Console.WriteLine("Show Recordset Information")
                    Console.WriteLine("#################")
                    If Me.RS Is Nothing Then
                        Console.WriteLine("Me.RS Is Nothing")
                    Else
                        With Me.RS
                            Console.WriteLine("EOF=" & .EOF)
                            If .EOF = False Then
                                Console.WriteLine("Fields.Count=" & .Fields.Count)
                                If .Fields.Count > 0 Then
                                    Dim i As Integer
                                    For i = 0 To .Fields.Count - 1
                                        Console.WriteLine(".Fields.Item(" & i & ").Name=" & .Fields.Item(i).Name & "[" & .Fields.Item(i).Value.ToString & "]")
                                        Console.WriteLine(".Fields.Item(" & i & ").TypeName=" & .Fields.Item(i).TypeName)
                                        Console.WriteLine(".Fields.Item(" & i & ").FieldTypeName=" & .Fields.Item(i).FieldTypeName)
                                        Console.WriteLine(".Fields.Item(" & i & ").DataCategory=" & .Fields.Item(i).DataCategory)
                                    Next
                                End If
                            End If
                        End With
                    End If
                Case ConsoleKey.F
                    Console.WriteLine("#################")
                    Console.WriteLine("Recordset.MoveNext")
                    Console.WriteLine("#################")
                    If Me.RS Is Nothing Then
                        Console.WriteLine("Me.RS Is Nothing")
                    Else
                        With Me.RS
                            .MoveNext()
                            If .LastErr <> "" Then
                                Console.WriteLine("MoveNext Error:" & .LastErr)
                            Else
                                Console.WriteLine("MoveNext OK")
                            End If
                        End With
                    End If
                Case ConsoleKey.G
                    Console.WriteLine("#################")
                    Console.WriteLine("Recordset.NextRecordset")
                    Console.WriteLine("#################")
                    Me.RS = Me.RS.NextRecordset
                    If Me.RS.LastErr <> "" Then
                        Console.WriteLine("Error:" & Me.RS.LastErr)
                    ElseIf Me.rs Is Nothing Then
                        Console.WriteLine("NextRecordset is nothing")
                    Else
                        Console.WriteLine("OK")
                        'Me.RS = Me.RS2
                        With Me.RS
                            If .LastErr <> "" Then
                                Console.WriteLine("Error:" & .LastErr)
                            Else
                                Console.WriteLine("EOF=" & .EOF)
                                If .EOF = False Then
                                    Console.WriteLine("Fields.Count=" & .Fields.Count)
                                    If .Fields.Count > 0 Then
                                        Dim i As Integer
                                        For i = 0 To .Fields.Count - 1
                                            Console.WriteLine(".Fields.Item(" & i & ").Name=" & .Fields.Item(i).Name & "[" & .Fields.Item(i).Value.ToString & "]")
                                        Next
                                    End If
                                End If
                            End If
                        End With
                    End If
                Case ConsoleKey.H
                    Console.WriteLine("#################")
                    Console.WriteLine("Test ExecuteNonQuery")
                    Console.WriteLine("#################")
                    Console.CursorVisible = True
                    Console.WriteLine("Input SQL:")
                    Me.SQL = Console.ReadLine()
                    Dim oCmdSQLSrvText As New CmdSQLSrvText(Me.SQL)
                    With oCmdSQLSrvText
                        .ActiveConnection = Me.ConnSQLSrv.Connection
                        .ExecuteNonQuery()
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                        End If
                        Console.WriteLine("RecordsAffected=" & .RecordsAffected)
                    End With
                    Console.WriteLine("Input SpName:")
                    Me.SQL = Console.ReadLine()
                    Dim oCmdSQLSrvSp As New CmdSQLSrvSp(Me.SQL)
                    With oCmdSQLSrvSp
                        .ActiveConnection = Me.ConnSQLSrv.Connection
                        .ExecuteNonQuery()
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                        End If
                        Console.WriteLine("RecordsAffected=" & .RecordsAffected)
                    End With
                Case ConsoleKey.I
                    If Me.ConnSQLSrv.IsDBConnReady = False Then
                        Console.WriteLine(" Connection is not ready please OpenOrKeepActive")
                    Else
                        Console.WriteLine("*******************")
                        Console.WriteLine("Test JSon")
                        Console.WriteLine("*******************")
                        Console.WriteLine("Press Q to Up")
                        Console.WriteLine("Press A to Convert current row to JSON")
                        Console.WriteLine("Press B to Convert current recordset to JSON")
                        Console.WriteLine("Press C to Convert all recordset to JSON")
                        Console.WriteLine("Press D to Convert current recordset to Simple JSON Array")
                        Do While True
                            Me.CurrConsoleKey = Console.ReadKey(True).Key
                            Select Case Me.CurrConsoleKey
                                Case ConsoleKey.Q
                                    Exit Do
                                Case Else
                                    Dim oConsoleKey As ConsoleKey = Me.CurrConsoleKey
                                    Console.WriteLine("Input SQL:")
                                    Dim strSQL As String = Console.ReadLine
                                    Dim oCmdSQLSrvText As CmdSQLSrvText = New CmdSQLSrvText(strSQL)
                                    If oCmdSQLSrvText.LastErr <> "" Then
                                        Console.WriteLine(oCmdSQLSrvText.LastErr)
                                    Else
                                        Select Case oConsoleKey
                                            Case ConsoleKey.A
                                                With oCmdSQLSrvText
                                                    .ActiveConnection = Me.ConnSQLSrv.Connection
                                                    Console.WriteLine("Execute")
                                                    Me.RS = .Execute()
                                                    If .LastErr <> "" Then
                                                        Console.WriteLine(.LastErr)
                                                    Else
                                                        Console.WriteLine("OK")
                                                        If Me.RS.EOF = True Then
                                                            Console.WriteLine("EOF=" & Me.RS.EOF)
                                                        End If
                                                        Console.WriteLine("Row2JSon=" & Me.RS.Row2JSon())
                                                    End If
                                                    Me.RS.Close()
                                                    Me.RS = Nothing
                                                    Exit Do
                                                End With
                                            Case ConsoleKey.B
                                                With oCmdSQLSrvText
                                                    .ActiveConnection = Me.ConnSQLSrv.Connection
                                                    Console.WriteLine("Execute")
                                                    Me.RS = .Execute()
                                                    If .LastErr <> "" Then
                                                        Console.WriteLine(.LastErr)
                                                    Else
                                                        Console.WriteLine("OK")
                                                        If Me.RS.EOF = True Then
                                                            Console.WriteLine("EOF=" & Me.RS.EOF)
                                                        End If
                                                        Console.WriteLine("Recordset2JSon=" & Me.RS.Recordset2JSon)
                                                    End If
                                                    Me.RS.Close()
                                                    Me.RS = Nothing
                                                    Exit Do
                                                End With
                                            Case ConsoleKey.C
                                                With oCmdSQLSrvText
                                                    .ActiveConnection = Me.ConnSQLSrv.Connection
                                                    Console.WriteLine("Execute")
                                                    Me.RS = .Execute()
                                                    If .LastErr <> "" Then
                                                        Console.WriteLine(.LastErr)
                                                    Else
                                                        Console.WriteLine("OK")
                                                        If Me.RS.EOF = True Then
                                                            Console.WriteLine("EOF=" & Me.RS.EOF)
                                                        End If
                                                        Console.WriteLine("AllRecordset2JSon=" & Me.RS.AllRecordset2JSon)
                                                    End If
                                                    Me.RS.Close()
                                                    Me.RS = Nothing
                                                    Exit Do
                                                End With
                                            Case ConsoleKey.D
                                                With oCmdSQLSrvText
                                                    .ActiveConnection = Me.ConnSQLSrv.Connection
                                                    Console.WriteLine("Execute")
                                                    Me.RS = .Execute()
                                                    If .LastErr <> "" Then
                                                        Console.WriteLine(.LastErr)
                                                    Else
                                                        Console.WriteLine("OK")
                                                        If Me.RS.EOF = True Then
                                                            Console.WriteLine("EOF=" & Me.RS.EOF)
                                                        End If
                                                        Console.WriteLine("Recordset2SimpleJSonArray=" & Me.RS.Recordset2SimpleJSonArray)
                                                    End If
                                                    Me.RS.Close()
                                                    Me.RS = Nothing
                                                    Exit Do
                                                End With
                                        End Select
                                    End If
                            End Select
                        Loop
                    End If
                Case ConsoleKey.J
                    If Me.ConnSQLSrv.IsDBConnReady = False Then
                        Console.WriteLine(" Connection is not ready please OpenOrKeepActive")
                    Else
                        Console.WriteLine("*******************")
                        Console.WriteLine("Execute SQL Server StoredProcedure")
                        Console.WriteLine("*******************")
                        Dim oCmdSQLSrvSp As New CmdSQLSrvSp("sp_helpdb")
                        With oCmdSQLSrvSp
                            .ActiveConnection = Me.ConnSQLSrv.Connection
                            .AddPara("@dbname", SqlDbType.NVarChar, 128)
                            '.ParaValue("@dbname") = "master"
                            Console.WriteLine("ParaValue(@dbname)=" & .ParaValue("@dbname"))
                            Console.WriteLine("DebugStr=" & .DebugStr)
                            Console.WriteLine("Execute")
                            Dim rsAny As Recordset = .Execute()
                            If .LastErr <> "" Then
                                Console.WriteLine(.LastErr)
                            Else
                                Console.WriteLine("OK")
                                Console.WriteLine("RecordsAffected=" & rsAny.RecordsAffected)
                                Console.WriteLine("ReturnValue=" & .ReturnValue)
                                With rsAny
                                    Console.WriteLine("Fields.Count=" & .Fields.Count)
                                    If .Fields.Count > 0 Then
                                        Dim i As Integer
                                        For i = 0 To .Fields.Count - 1
                                            Console.WriteLine(".Fields.Item(" & i & ").Name=" & .Fields.Item(i).Name & "[" & .Fields.Item(i).Value.ToString & "]")
                                        Next
                                    End If
                                    Console.WriteLine("EOF=" & .EOF)
                                End With
                            End If
                            rsAny.Close()
                            rsAny = Nothing
                        End With
                    End If
                Case ConsoleKey.K
                    If Me.ConnSQLSrv.IsDBConnReady = False Then
                        Console.WriteLine(" Connection is not ready please OpenOrKeepActive")
                    Else
                        Console.WriteLine("*******************")
                        Console.WriteLine("Execute SQL Server SQL statement Text")
                        Console.WriteLine("*******************")
                        Dim oCmdSQLSrvText As New CmdSQLSrvText("select * from master.dbo.sysdatabases where name = @name")
                        With oCmdSQLSrvText
                            .ActiveConnection = Me.ConnSQLSrv.Connection
                            .AddPara("@name", SqlDbType.VarChar, 128)
                            .ParaValue("@name") = "master"
                            Console.WriteLine("ParaValue(@name)=" & .ParaValue("@name"))
                            Console.WriteLine("DebugStr=" & .DebugStr)
                            Console.WriteLine("Execute")
                            Dim oRS As Recordset = .Execute()
                            If .LastErr <> "" Then
                                Console.WriteLine(.LastErr)
                            Else
                                Console.WriteLine("OK")
                                Console.WriteLine("RecordsAffected=" & oRS.RecordsAffected)
                                With oRS
                                    Console.WriteLine("Fields.Count=" & .Fields.Count)
                                    If .Fields.Count > 0 Then
                                        Dim i As Integer
                                        For i = 0 To .Fields.Count - 1
                                            Console.WriteLine(".Fields(" & i & ").Name=" & .Fields.Item(i).Name & "[" & .Fields.Item(i).Value.ToString & "]")
                                        Next
                                    End If
                                    Console.WriteLine("EOF=" & .EOF)
                                End With
                            End If
                            oRS.Close()
                            oRS = Nothing
                        End With
                    End If
                Case ConsoleKey.L
                    If Me.ConnSQLSrv Is Nothing Then
                        Console.WriteLine("ConnSQLSrv Is Nothing")
                    ElseIf Me.ConnSQLSrv.IsDBConnReady = False Then
                        Console.WriteLine("Me.ConnSQLSrv.IsDBConnReady = False")
                    Else
                        Dim oSQLSrvTools As New SQLSrvTools(Me.ConnSQLSrv)
                        With oSQLSrvTools
                            Console.WriteLine(".IsDatabaseExists(master)=" & .IsDatabaseExists("master"))
                            If .LastErr <> "" Then Console.WriteLine(.LastErr)
                            Console.WriteLine("Input TabName")
                            Dim strTabName As String = Console.ReadLine
                            Console.WriteLine(".IsDatabaseExists(" & strTabName & ")=" & .IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, strTabName))
                            If .LastErr <> "" Then Console.WriteLine(.LastErr)
                            Console.WriteLine("Input DBUser")
                            Dim strDBUser As String = Console.ReadLine
                            Console.WriteLine(".IsDatabaseExists(" & strDBUser & ")=" & .IsDBUserExists(strDBUser))
                            If .LastErr <> "" Then Console.WriteLine(.LastErr)
                            Console.WriteLine(".IsLoginUserExists(sa)=" & .IsLoginUserExists("sa"))
                            If .LastErr <> "" Then Console.WriteLine(.LastErr)
                            Console.WriteLine("Input TableName=" & Me.TableName)
                            Me.TableName = Console.ReadLine()
                            Console.WriteLine("Input ColName=" & Me.ColName)
                            Me.ColName = Console.ReadLine()
                            Console.WriteLine(".IsTabColExists(" & Me.TableName & "，" & Me.ColName & ")=" & .IsTabColExists(Me.TableName, Me.ColName))
                            If .LastErr <> "" Then Console.WriteLine(.LastErr)
                        End With
                    End If
                Case ConsoleKey.M
                    If Me.ConnSQLSrv Is Nothing Then
                        Console.WriteLine("ConnSQLSrv Is Nothing")
                    ElseIf Me.ConnSQLSrv.IsDBConnReady = False Then
                        Console.WriteLine("ConnSQLSrv.IsDBConnReady=" & Me.ConnSQLSrv.IsDBConnReady)
                    Else
                        Dim oCmdSQLSrvText1 As New CmdSQLSrvText("select * from sysusers")
                        Dim oCmdSQLSrvText2 As New CmdSQLSrvText("select * from sysobjects")
                        Dim oCmdSQLSrvText3 As New CmdSQLSrvText("create table t1(f1 int)")
                        Dim rs1 As Recordset, rs2 As Recordset
                        With oCmdSQLSrvText3
                            .ActiveConnection = Me.ConnSQLSrv.Connection
                            Console.WriteLine("oCmdSQLSrvText3.ExecuteNonQuery")
                            .ExecuteNonQuery()
                            If .LastErr <> "" Then
                                Console.WriteLine(.LastErr)
                            Else
                                Console.WriteLine("OK")
                                Console.WriteLine("RecordsAffected=" & .RecordsAffected)
                            End If
                        End With
                        With oCmdSQLSrvText1
                            .ActiveConnection = Me.ConnSQLSrv.Connection
                            Console.WriteLine("oCmdSQLSrvText1.Execute")
                            rs1 = .Execute
                            If .LastErr <> "" Then
                                Console.WriteLine(.LastErr)
                            Else
                                Console.WriteLine("OK")
                                With rs1
                                    Console.WriteLine("EOF=" & .EOF)
                                    Console.WriteLine("Row2JSon=" & .Row2JSon)
                                End With
                            End If
                        End With
                        With oCmdSQLSrvText2
                            .ActiveConnection = Me.ConnSQLSrv.Connection
                            Console.WriteLine("oCmdSQLSrvText2.Execute")
                            rs2 = .Execute
                            If .LastErr <> "" Then
                                Console.WriteLine(.LastErr)
                            Else
                                Console.WriteLine("OK")
                                With rs2
                                    Console.WriteLine("EOF=" & .EOF)
                                    Console.WriteLine("Row2JSon=" & .Row2JSon)
                                End With
                            End If
                        End With
                    End If
                Case ConsoleKey.N
                    Console.WriteLine("*******************")
                    Console.WriteLine("Test Cache Query")
                    Console.WriteLine("*******************")
                    If Me.ConnSQLSrv Is Nothing Then
                        Console.WriteLine("ConnSQLSrv Is Nothing")
                    ElseIf Me.ConnSQLSrv.IsDBConnReady = False Then
                        Console.WriteLine("IsDBConnReady = False")
                    Else
                        Console.WriteLine("ConnSQLSrv.IsDBConnReady=" & Me.ConnSQLSrv.IsDBConnReady)
                        Console.WriteLine("Press Q to Up")
                        Console.WriteLine("Press A to CmdSQLSrvText")
                        Console.WriteLine("Press B to CmdSQLSrvSp")
                        Do While True
                            Me.CurrConsoleKey = Console.ReadKey(True).Key
                            Select Case Me.CurrConsoleKey
                                Case ConsoleKey.Q
                                    Exit Do
                                Case ConsoleKey.A
                                    Console.CursorVisible = True
                                    Dim oCmdSQLSrvText As New CmdSQLSrvText("select * from sysobjects where name=@name")
                                    'oCmdSQLSrvText.ActiveConnection = Me.ConnSQLSrv.Connection
                                    oCmdSQLSrvText.AddPara("@name", SqlDbType.VarChar, 256)
                                    Console.WriteLine("Input db object name=sysobjects")
                                    Dim strName As String = Console.ReadLine()
                                    If strName = "" Then strName = "sysobjects"
                                    oCmdSQLSrvText.ParaValue("@name") = strName
                                    Dim strKeyName As String = oCmdSQLSrvText.KeyName
                                    Console.WriteLine("InitPigKeyValue=")
                                    Me.ConnSQLSrv.InitPigKeyValue()
                                    Console.WriteLine(Me.ConnSQLSrv.LastErr)
                                    'Console.WriteLine("Before IsPigKeyValueExists(" & strKeyName & ")=" & Me.ConnSQLSrv.PigKeyValueApp.IsPigKeyValueExists(strKeyName))
                                    Console.WriteLine("CacheQuery=")
                                    Dim strJSon As String = oCmdSQLSrvText.CacheQuery(Me.ConnSQLSrv)
                                    Console.WriteLine(oCmdSQLSrvText.LastErr)
                                    'Console.WriteLine("After IsPigKeyValueExists(" & strKeyName & ")=" & Me.ConnSQLSrv.PigKeyValueApp.IsPigKeyValueExists(strKeyName))
                                    Console.WriteLine("JSon=" & strJSon)
                                    Exit Do
                                Case ConsoleKey.B
                                    Console.CursorVisible = True
                                    Dim oCmdSQLSrvSp As New CmdSQLSrvSp("sp_helpdb")
                                    'oCmdSQLSrvSp.ActiveConnection = Me.ConnSQLSrv.Connection
                                    oCmdSQLSrvSp.AddPara("@dbname", SqlDbType.VarChar, 256)
                                    oCmdSQLSrvSp.ParaValue("@dbname") = "master"
                                    Dim strKeyName As String = oCmdSQLSrvSp.KeyName
                                    Console.WriteLine("InitPigKeyValue=")
                                    Me.ConnSQLSrv.InitPigKeyValue()
                                    Console.WriteLine(Me.ConnSQLSrv.LastErr)
                                    'Console.WriteLine("Before IsPigKeyValueExists(" & strKeyName & ")=" & Me.ConnSQLSrv.PigKeyValueApp.IsPigKeyValueExists(strKeyName))
                                    Console.WriteLine("CacheQuery=")
                                    Dim strJSon As String = oCmdSQLSrvSp.CacheQuery(Me.ConnSQLSrv)
                                    Console.WriteLine(oCmdSQLSrvSp.LastErr)
                                    'Console.WriteLine("After IsPigKeyValueExists(" & strKeyName & ")=" & Me.ConnSQLSrv.PigKeyValueApp.IsPigKeyValueExists(strKeyName))
                                    Console.WriteLine("JSon=" & strJSon)
                                    Exit Do
                            End Select
                        Loop
                    End If
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub


    'Public Sub ShowDBConnDef(ByRef oDBConnDef As DBConnDef)
    '    With oDBConnDef
    '        Console.WriteLine("DBConnName=" & .DBConnName)
    '        Console.WriteLine("CurrDatabase=" & .CurrDatabase)
    '        Console.WriteLine("DBUser=" & .DBUser)
    '        Console.WriteLine("DBUserPwd=" & .DBUserPwd)
    '    End With
    'End Sub

    Public Sub MainSet()
        Do While True
            Console.Clear()
            Me.MenuDefinition = "GenerateEncKey#Generate EncKey|"
            Me.MenuDefinition &= "NewDBConnMgr#New DBConnMgr|"
            Me.MenuDefinition &= "ShowDBConnDefs#Show DBConnDefs|"
            Me.MenuDefinition &= "LoadDBConnDefs#LoadDBConnDefs|"
            Me.MenuDefinition &= "AddOrGetDBConnDef#AddOrGet DBConnDef|"
            Me.MenuDefinition &= "SaveDBConnDefs#SaveDBConnDefs|"
            Me.MenuDefinition &= "EditDBConnDef#Edit DBConnDef|"
            Me.MenuDefinition &= "RemoveDBConnDef#Remove DBConnDef|"
            Me.PigConsole.SimpleMenu("Database connection management", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "ShowDBConnDefs"
                    'If Me.DBConnMgr Is Nothing Then
                    '    Console.WriteLine("Me.DBConnMgr Is Nothing")
                    'Else
                    '    Console.WriteLine("DBConnDefs.Count=" & Me.DBConnMgr.DBConnDefs.Count)
                    '    For Each oDBConnDef As DBConnDef In Me.DBConnMgr.DBConnDefs
                    '        Me.ShowDBConnDef(oDBConnDef)
                    '    Next
                    'End If
                Case "GenerateEncKey"
                    'Me.DBConnMgr = New DBConnMgr("")
                    'Me.Ret = Me.DBConnMgr.MkEncKey(Me.EncKey)
                    'If Me.Ret <> "OK" Then
                    '    Console.WriteLine("MkEncKey=" & Me.Ret)
                    'Else
                    '    Console.WriteLine("EncKey=")
                    '    Console.WriteLine(Me.EncKey)
                    'End If
                Case "NewDBConnMgr"
                    'Me.PigConsole.GetLine("Input ConfFilePath", Me.ConfFilePath)
                    'Me.PigConsole.GetLine("Input MkEncKey", Me.EncKey)
                    'Me.DBConnMgr = New DBConnMgr(Me.EncKey, Me.ConfFilePath)
                    'Console.WriteLine("New DBConnMgr...")
                    'If Me.DBConnMgr.LastErr <> "" Then
                    '    Console.WriteLine(Me.DBConnMgr.LastErr)
                    'Else
                    '    Console.WriteLine("OK")
                    'End If
                Case "LoadDBConnDefs"
                    'If Me.DBConnMgr Is Nothing Then
                    '    Console.WriteLine("Me.DBConnMgr Is Nothing")
                    'Else
                    '    Console.WriteLine("LoadDBConnDefs...")
                    '    Me.Ret = Me.DBConnMgr.LoadDBConnDefs()
                    '    Console.WriteLine(Me.Ret)
                    'End If
                Case "AddOrGetDBConnDef"
                    'Me.PigConsole.GetLine("Input DBConnName", Me.DBConnName)
                    'Me.PigConsole.GetLine("Input SQLServer", Me.DBSrv)
                    'Me.PigConsole.GetLine("Input CurrDatabase", Me.CurrDB)
                    'Me.PigConsole.GetLine("Input DBUser", Me.DBUser)
                    'Me.DBPwd = Me.PigConsole.GetPwdStr("Input DBPwd")
                    'If Me.DBUser = "" Then
                    '    Me.DBConnMgr.DBConnDefs.Add(Me.DBConnName, Me.DBSrv, Me.CurrDB)
                    'Else
                    '    Me.DBConnMgr.DBConnDefs.Add(Me.DBConnName, Me.DBSrv, Me.CurrDB, Me.DBUser, Me.DBPwd)
                    'End If
                    'If Me.DBConnMgr.LastErr <> "" Then
                    '    Console.WriteLine("New DBConnMgr" & Me.DBConnMgr.LastErr)
                    'End If
                Case "SaveDBConnDefs"
                    'If Me.DBConnMgr Is Nothing Then
                    '    Console.WriteLine("Me.DBConnMgr Is Nothing")
                    'Else
                    '    Me.Ret = Me.DBConnMgr.SaveDBConnDefs()
                    '    Console.WriteLine("SaveDBConnDefs=" & Me.Ret)
                    'End If
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub

    Public Sub Main()
        Do While True
            Console.Clear()
            Me.MenuDefinition = ""
            'Me.MenuDefinition &= "MainSet#Database connection management|"
            Me.MenuDefinition &= "MainFunc#Main Function Demo|"
            Me.MenuDefinition &= "SQLSrvTools#SQLSrvTools Demo|"
            Me.PigConsole.SimpleMenu("Main Menu", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoExit)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "SQLSrvTools"
                    Me.SQLSrvToolsDemo()
                Case "MainSet"
                    Me.MainSet()
                Case "MainFunc"
                    Me.MainFunc()
            End Select
        Loop
    End Sub

    Public Sub SQLSrvToolsDemo()
        Do While True
            Console.Clear()
            Me.MenuDefinition = ""
            Me.MenuDefinition &= "GetTableOrView2VBCode#Generate VB class code corresponding to table or view|"
            Me.MenuDefinition &= "GetTableOrView2SQLOrVBFragment#Generate SQL statement or VB code fragments corresponding to tables or views|"
            Me.PigConsole.SimpleMenu("SQLSrvToolsDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "GetTableOrView2VBCode", "GetTableOrView2SQLOrVBFragment"
                    If Me.SQLSrvTools Is Nothing Then
                        If Me.ConnSQLSrv Is Nothing Then
                            Console.WriteLine("ConnSQLSrv Is Nothing")
                        ElseIf Me.ConnSQLSrv.IsDBConnReady = False Then
                            Console.WriteLine("ConnSQLSrv.IsDBConnReady = False")
                        Else
                            Me.SQLSrvTools = New SQLSrvTools(Me.ConnSQLSrv)
                        End If
                    End If
                    If Me.SQLSrvTools IsNot Nothing Then
                        Select Case Me.MenuKey
                            Case "GetTableOrView2VBCode"
                                Me.PigConsole.GetLine("Input table or view name", Me.TableName)
                                Me.PigConsole.GetLine("Input NotMathFillByRsList ,separated by ','", Me.NotMathFillByRsList)
                                Me.PigConsole.GetLine("Input NotMathMD5List ,separated by ','", Me.NotMathMD5List)
                                Me.PigConsole.GetLine("Input save filepath", Me.FilePath)
                                Dim bolIsSimpleProperty As Boolean = Me.PigConsole.IsYesOrNo("Is Simple Property")
                                Dim bolIsSetUpdateTime As Boolean
                                If bolIsSimpleProperty = False Then
                                    bolIsSetUpdateTime = Me.PigConsole.IsYesOrNo("Is Set UpdateTime")
                                End If
                                Me.Ret = Me.SQLSrvTools.GetTableOrView2VBCode(Me.TableName, Me.VBCodeOrSQLFragment, Me.NotMathFillByRsList, Me.NotMathMD5List, bolIsSimpleProperty, bolIsSetUpdateTime)
                                If Me.Ret <> "OK" Then
                                    Console.WriteLine(Me.Ret)
                                Else
                                    Console.WriteLine("Save to " & Me.FilePath)
                                    Me.PigFunc.SaveTextToFile(Me.FilePath, Me.VBCodeOrSQLFragment)
                                End If
                            Case "GetTableOrView2SQLOrVBFragment"
                                Me.PigConsole.GetLine("Input table or view name", Me.TableName)
                                Me.PigConsole.GetLine("List of unwanted column names, separated by ','", Me.NotMathColList)
                                Me.MenuDefinition2 = ""
                                Me.MenuDefinition2 &= CStr(SQLSrvTools.EnmWhatFragment.SpInParas) & "#" & SQLSrvTools.EnmWhatFragment.SpInParas.ToString & "|"
                                Me.MenuDefinition2 &= CStr(SQLSrvTools.EnmWhatFragment.SpInParasSetNull) & "#" & SQLSrvTools.EnmWhatFragment.SpInParasSetNull.ToString & "|"
                                Me.MenuDefinition2 &= CStr(SQLSrvTools.EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_AddPara) & "#" & SQLSrvTools.EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_AddPara.ToString & "|"
                                Me.MenuDefinition2 &= CStr(SQLSrvTools.EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_ParaValue) & "#" & SQLSrvTools.EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_ParaValue.ToString & "|"
                                Me.MenuDefinition2 &= CStr(SQLSrvTools.EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_AddPara_ParaValue) & "#" & SQLSrvTools.EnmWhatFragment.CmdSQLSrvSpOrCmdSQLSrvText_AddPara_ParaValue.ToString & "|"
                                Me.MenuDefinition2 &= CStr(SQLSrvTools.EnmWhatFragment.UpdatePerCol) & "#" & SQLSrvTools.EnmWhatFragment.UpdatePerCol.ToString & "|"
                                Me.PigConsole.SimpleMenu("Select WhatFragment", Me.MenuDefinition2, Me.MenuKey2, PigConsole.EnmSimpleMenuExitType.Null)
                                Me.WhatFragment = CInt(Me.MenuKey2)
                                Me.PigConsole.GetLine("Input save filepath", Me.FilePath)
                                Me.Ret = Me.SQLSrvTools.GetTableOrView2SQLOrVBFragment(Me.TableName, Me.WhatFragment, Me.VBCodeOrSQLFragment, Me.NotMathColList)
                                If Me.Ret <> "OK" Then
                                    Console.WriteLine(Me.Ret)
                                Else
                                    Console.WriteLine("Save to " & Me.FilePath)
                                    Me.PigFunc.SaveTextToFile(Me.FilePath, Me.VBCodeOrSQLFragment)
                                End If
                        End Select
                    End If
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
