Imports System.Data
Imports PigSQLSrvLib

Public Class ConsoleDemo
    Public ConnSQLSrv As ConnSQLSrv
    Public ConnStr As String
    Public SQL As String
    Public RS As Recordset
    Public DBSrv As String = "localhost"
    Public MirrDBSrv As String = "localhost"
    Public DBUser As String = "sa"
    Public DBPwd As String = ""
    Public CurrDB As String = "master"
    Public CurrConsoleKey As ConsoleKey
    Public InpStr As String
    Public AccessFilePath As String

    Public Sub Main()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to Set SQL Server Connection String")
            Console.WriteLine("Press B to OpenOrKeepActive Connection")
            Console.WriteLine("Press C to Show Connection Information")
            'Console.WriteLine("Press D to Create Recordset with Execute")
            Console.WriteLine("Press E to Show Recordset Information")
            Console.WriteLine("Press F to Recordset.MoveNext")
            'Console.WriteLine("Press G to Recordset.NextRecordset")
            'Console.WriteLine("Press H to Test Command")
            'Console.WriteLine("Press I to Test JSon")
            Console.WriteLine("Press J to Execute SQL Server StoredProcedure")
            Console.WriteLine("Press K to Execute SQL Server SQL statement Text")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey().Key
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
                        Me.CurrConsoleKey = Console.ReadKey().Key
                        Select Case Me.CurrConsoleKey
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Console.WriteLine("Input SQL Server:" & Me.DBSrv)
                                Me.DBSrv = Console.ReadLine()
                                If Me.DBSrv = "" Then Me.DBSrv = "localhost"
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
                                        Me.DBPwd = Console.ReadLine()
                                        Console.WriteLine("DB Password=" & Me.DBPwd)
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
                                        Me.DBPwd = Console.ReadLine()
                                        Console.WriteLine("DB Password=" & Me.DBPwd)
                                        Me.ConnSQLSrv = New ConnSQLSrv(Me.DBSrv, Me.MirrDBSrv, Me.CurrDB, Me.DBUser, Me.DBPwd)
                                End Select
                                Exit Do
                        End Select
                    Loop
                Case ConsoleKey.B
                    Console.WriteLine("#################")
                    Console.WriteLine("Open Connection")
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
                    'Console.WriteLine("#################")
                    'Console.WriteLine("Create Recordset with Execute")
                    'Console.WriteLine("#################")
                    'Console.WriteLine("Input SQL:")
                    'Me.SQL = Console.ReadLine()
                    'With Me.ConnSQLSrv
                    '    Console.WriteLine("Execute:")
                    '    Me.RS = .Execute(SQL)
                    '    If .LastErr <> "" Then
                    '        Console.WriteLine(.LastErr)
                    '    Else
                    '        Console.WriteLine("OK")
                    '    End If
                    'End With
                Case ConsoleKey.E
                    Console.WriteLine("#################")
                    Console.WriteLine("Show Recordset Information")
                    Console.WriteLine("#################")
                    With Me.RS
                        Console.WriteLine("Fields.Count=" & .Fields.Count)
                        If .Fields.Count > 0 Then
                            Dim i As Integer
                            For i = 0 To .Fields.Count - 1
                                Console.WriteLine(".Fields.Item(" & i & ").Name=" & .Fields.Item(i).Name & "[" & .Fields.Item(i).Value.ToString & "]")
                                Console.WriteLine(".Fields.Item(" & i & ").Type=" & .Fields.Item(i).Type.ToString)
                            Next
                        End If
                        Console.WriteLine("EOF=" & .EOF)
                    End With
                Case ConsoleKey.F
                    Console.WriteLine("#################")
                    Console.WriteLine("Recordset.MoveNext")
                    Console.WriteLine("#################")
                    With Me.RS
                        .MoveNext()
                        If .LastErr <> "" Then
                            Console.WriteLine("MoveNext Error:" & .LastErr)
                        Else
                            Console.WriteLine("MoveNext OK")
                        End If
                    End With
                Case ConsoleKey.G
                    'Console.WriteLine("#################")
                    'Console.WriteLine("Recordset.NextRecordset")
                    'Console.WriteLine("#################")
                    'Me.RS = Me.RS.NextRecordset
                    'With Me.RS
                    '    'Dim oRs As Recordset = .NextRecordset
                    '    If .LastErr <> "" Then
                    '        Console.WriteLine("Error:" & .LastErr)
                    '    Else
                    '        Console.WriteLine("OK")
                    '        Console.WriteLine("Fields.Count=" & .Fields.Count)
                    '        If .Fields.Count > 0 Then
                    '            Dim i As Integer
                    '            For i = 0 To .Fields.Count - 1
                    '                Console.WriteLine(".Fields.Item(" & i & ").Name=" & .Fields.Item(i).Name & "[" & .Fields.Item(i).Value.ToString & "]")
                    '            Next
                    '        End If
                    '        Console.WriteLine("EOF=" & .EOF)
                    '    End If
                    'End With
                'Case ConsoleKey.H
                '    Console.WriteLine("#################")
                '    Console.WriteLine("Test Command")
                '    Console.WriteLine("#################")
                '    Dim oCommand As New Command
                '    With oCommand
                '        Console.WriteLine("Set ActiveConnection")
                '        .ActiveConnection = Me.ConnSQLSrv.Connection
                '        Console.WriteLine("CommandText=""select * from master.dbo.sysdatabases where name = ?")
                '        .CommandText = "select * from master.dbo.sysdatabases where name = ?"
                '        Console.WriteLine("CreateParameter @dbname=""master""")
                '        .Parameters.Append(.CreateParameter("@dbname", Field.DataTypeEnum.adVarChar, Parameter.ParameterDirectionEnum.adParamInput, 128, "master"))
                '        .Parameters.Item("@dbname").Value = "WxWorkDB"
                '        Console.WriteLine("Parameters.Item(@dbname).Value=" & .Parameters.Item("@dbname").Value)
                '        If .LastErr <> "" Then
                '            Console.WriteLine(.LastErr)
                '        Else
                '            Console.WriteLine("OK")
                '        End If
                '        Console.WriteLine("Execute")
                '        Dim rsAny = .Execute()
                '        If .LastErr <> "" Then
                '            Console.WriteLine(.LastErr)
                '        Else
                '            Console.WriteLine("OK")
                '            With rsAny
                '                Console.WriteLine("Fields.Count=" & .Fields.Count)
                '                If .Fields.Count > 0 Then
                '                    Dim i As Integer
                '                    For i = 0 To .Fields.Count - 1
                '                        Console.WriteLine(".Fields.Item(" & i & ").Name=" & .Fields.Item(i).Name & "[" & .Fields.Item(i).Value.ToString & "]")
                '                    Next
                '                End If
                '                Console.WriteLine("PageCount=" & .PageCount)
                '                Console.WriteLine("EOF=" & .EOF)
                '            End With
                '        End If
                '        .Parameters.Delete("@dbname")
                '    End With
                'Case ConsoleKey.I
                '    Console.WriteLine("*******************")
                '    Console.WriteLine("Test JSon")
                '    Console.WriteLine("*******************")
                '    Console.WriteLine("Press Q to Up")
                '    Console.WriteLine("Press A to Convert current row to JSON")
                '    Console.WriteLine("Press B to Convert current recordset to JSON")
                '    Console.WriteLine("Press C to Convert all recordset to JSON")
                '    Do While True
                '        Me.CurrConsoleKey = Console.ReadKey().Key
                '        Select Case Me.CurrConsoleKey
                '            Case ConsoleKey.Q
                '                Exit Do
                '            Case ConsoleKey.A
                '                Console.WriteLine(Me.RS.Row2JSon)
                '                If Me.RS.LastErr <> "" Then Console.WriteLine(Me.RS.LastErr)
                '                Exit Do
                '            Case ConsoleKey.B
                '                Console.WriteLine(Me.RS.Recordset2JSon(10))
                '                If Me.RS.LastErr <> "" Then Console.WriteLine(Me.RS.LastErr)
                '                Exit Do
                '            Case ConsoleKey.C
                '                Console.WriteLine(Me.RS.AllRecordset2JSon())
                '                If Me.RS.LastErr <> "" Then Console.WriteLine(Me.RS.LastErr)
                '                Exit Do
                '        End Select
                '    Loop
                Case ConsoleKey.J
                    Console.WriteLine("*******************")
                    Console.WriteLine("Execute SQL Server StoredProcedure")
                    Console.WriteLine("*******************")
                    Dim oCmdSQLSrvSp As New CmdSQLSrvSp("sp_helpdb")
                    With oCmdSQLSrvSp
                        .ActiveConnection = Me.ConnSQLSrv.Connection
                        .AddPara("@dbname", SqlDbType.NVarChar, 128)
                        .ParaValue("@dbname") = "master"
                        Console.WriteLine("ParaValue(@dbname)=" & .ParaValue("@dbname"))
                        Console.WriteLine("Execute")
                        Dim rsAny As Recordset = .Execute()
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                            Console.WriteLine("RecordsAffected=" & .RecordsAffected)
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
                Case ConsoleKey.K
                    Console.WriteLine("*******************")
                    Console.WriteLine("Execute SQL Server SQL statement Text")
                    Console.WriteLine("*******************")
                    Dim oCmdSQLSrvText As New CmdSQLSrvText("select * from master.dbo.sysdatabases where name = @name")
                    With oCmdSQLSrvText
                        .ActiveConnection = Me.ConnSQLSrv.Connection
                        .AddPara("@name", SqlDbType.VarChar, 128)
                        .ParaValue("@name") = "master"
                        Console.WriteLine("ParaValue(@name)=" & .ParaValue("@name"))
                        Console.WriteLine("Execute")
                        Dim oRS As Recordset = .Execute()
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                            Console.WriteLine("RecordsAffected=" & .RecordsAffected)
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
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
