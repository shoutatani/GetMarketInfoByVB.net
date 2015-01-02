Imports System.Data.SQLite

Public Class SQLiteDB

    Private _con As SQLiteConnection

    ''' <summary>
    ''' open SQLite database in current folder with the database name 
    ''' </summary>
    ''' <param name="aDBName">databasename</param>
    ''' <returns>success or failure</returns>
    Public Function openSQLiteDatabase(ByVal aDBName As String) As Boolean
        Try

            If IsNothing(_con) = True Then
                _con = New SQLiteConnection
            End If

            _con.ConnectionString = "Data Source=" & System.IO.Directory.GetCurrentDirectory & "/" & aDBName & _
                                    ";Default Timeout=600" & _
                                    ";Synchronous=Full"
            If _con.State <> ConnectionState.Open Then
                _con.Open()
            End If

            Return True
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' close database
    ''' </summary>
    ''' <returns>success or failure</returns>
    Public Function closeSQLiteDatabase() As Boolean
        Try
            If IsNothing(_con) = False AndAlso _con.State <> ConnectionState.Closed Then
                _con.Close()
            End If
            Return True
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' check the table,if not exist and create command is defined,create the table
    ''' </summary>
    ''' <param name="aTableName">table name</param>
    ''' <param name="aCreateCmd">create command</param>
    ''' <returns>success or failure</returns>
    Public Function IsExistTable(ByVal aTableName As String, Optional ByVal aCreateCmd As String = "") As Boolean
        Try
            Dim _cmd As New SQLiteCommand
            Dim cntResult As Integer
            IsExistTable = False

            _cmd.Connection = _con

            _cmd.CommandText = "Select count(*) from sqlite_master" & _
                                " where type = 'table'" & _
                                " and tbl_name = '" & aTableName & "'"

            cntResult = CType(_cmd.ExecuteScalar(), Integer)

            If cntResult = 0 AndAlso aCreateCmd <> "" Then
                _cmd.CommandText = aCreateCmd
                If _cmd.ExecuteNonQuery() <> -1 Then
                    Return True
                Else
                    Return False
                End If
            ElseIf cntResult = 0 Then
                Return False
            End If
            Return True
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' update the table with the datatable
    ''' </summary>
    ''' <param name="aTableName">table name</param>
    ''' <param name="_dtCalData">target datatable</param>
    ''' <returns>success or failure</returns>
    Public Function UpdateDataTable(ByVal aTableName As String, ByVal _dtCalData As DataTable) As Boolean
        Try
            Dim _cmd As SQLiteCommand
            Dim _da As SQLiteDataAdapter
            Dim _dt As New DataTable
            Dim _cmdbd As SQLiteCommandBuilder

            _cmd = New SQLiteCommand(_con)
            _cmd.CommandText = "Select * from " & aTableName

            _da = New SQLiteDataAdapter(_cmd)
            _da.Fill(_dt)
            _cmdbd = New SQLiteCommandBuilder(_da)

            _dt.Merge(_dtCalData)
            _da.Update(_dt)

            Return True
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return False
        End Try
    End Function

End Class
