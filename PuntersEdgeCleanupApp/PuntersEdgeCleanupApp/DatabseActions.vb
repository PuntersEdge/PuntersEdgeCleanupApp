Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data.SqlClient

Public Module GlobalVars

    Public con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("PuntersEdgeDB").ConnectionString)
    Public command As New SqlCommand

End Module

Public Class DatabseActions

    Public Sub UPDATESession(ByVal table As String, value As String)

        Dim result As String = ""

        Try
            command.CommandType = CommandType.Text
            command.CommandText = "UPDATE [" & table & "] SET [SessionToken] = '" & value & "' WHERE [SessionToken] = '-1'"
            command.Connection = con
            con.Open()
            command.ExecuteNonQuery()
            con.Close()
            result = "Success:" & command.CommandText.ToString

        Catch ex As Exception
            result = ex.Message.ToString
        End Try

    End Sub 'For PuntersEdge ONLY - deals with session token for API access
    Public Sub INSERT(ByVal table As String, ByVal columns As String, values As String)


        command.CommandText = "INSERT INTO " & table & "(" & columns & ")" & " VALUES (" & values & ")"
        command.Connection = con

        con.OpenAsync()
        command.ExecuteNonQuery()
        con.Close()


    End Sub 'General INSERT statement. Accepts table name, columns in string (column1, column2 etc), and values in string ('varchar', int etc)
    Public Function SELECTSTATEMENT(ByVal ColumnsToSelect As String, ByVal table As String, ByVal whereClause As String)



        Dim strSql As String = "SELECT " & ColumnsToSelect & " FROM " & table & " " & whereClause
        Dim Resultset As New DataTable
        Using cnn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("PuntersEdgeDB").ConnectionString)
            cnn.Open()
            Using dad As New SqlDataAdapter(strSql, cnn)
                dad.Fill(Resultset)
            End Using
            cnn.Close()
        End Using


        Return Resultset
    End Function 'General Select statement. Accepts table name, columns or *, and a where clause if needed (if not, pass in "")
    Public Function SELECTSTATEMENT_Scalar(ByVal ColumnsToSelect As String, ByVal table As String, ByVal whereClause As String)



        Dim strSql As String = "SELECT " & ColumnsToSelect & " FROM " & table & " " & whereClause
        Dim Resultset As String
        Using cnn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("PuntersEdgeDB").ConnectionString)
            cnn.Open()

            command.CommandText = strSql
            command.Connection = cnn

            Resultset = command.ExecuteScalar

            cnn.Close()
        End Using


        Return Resultset
    End Function 'General Select statement. Accepts table name, columns or *, and a where clause if needed (if not, pass in "")
    Public Function EXECSPROC(ByVal SPROCNAME As String, variables As String)
        Dim Result As String = ""

        command.CommandText = "EXECUTE " & SPROCNAME & " " & variables
        command.Connection = con

        Try

            con.Open()
            Result = command.ExecuteScalar
            con.Close()

        Catch ex As Exception

        End Try

        Return Result

    End Function

    Public Sub UPDATE(ByVal table As String, columnToUpdate As String, valueToInsert As String, WHERE_CLAUSE As String)



        command.CommandText = "UPDATE " & table & " SET " & columnToUpdate & "='" & valueToInsert & "' " & WHERE_CLAUSE
        command.Connection = con

        con.Open()
        command.ExecuteNonQuery()
        con.Close()



    End Sub
    Public Sub SELECT_INTO(ByVal Table As String, ByVal columns As String, ByVal FromSQL As String)

        Using connection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("PuntersEdgeDB").ConnectionString)

            command.CommandText = "SELECT " & columns & " INTO " & Table & " FROM (" & FromSQL & ")"
            command.Connection = connection

            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()

        End Using





    End Sub

    Public Sub SQL(ByVal sql As String)

        command.CommandText = sql
        command.Connection = con

        con.Open()
        command.ExecuteNonQuery()
        con.Close()

    End Sub

End Class
