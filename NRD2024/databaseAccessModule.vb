Imports System.Data.OleDb
Imports System.Windows.Forms
' No need for Microsoft.Office.Interop.Excel unless it's used elsewhere in the module

Module DatabaseAccessModule
    Private connectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Application.StartupPath}\newNonReportableItems.accdb"

    Public Function OpenConnection() As OleDbConnection
        Dim connection As New OleDbConnection(connectionString)
        Try
            connection.Open()
            'MessageBox.Show("Connection Successful")
            Return connection
        Catch ex As Exception
            MessageBox.Show($"Error connecting to database: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ' Additional methods for database operations can be added here
    Public Function GetNonReportableItems() As DataTable
        Using connection As New OleDbConnection(connectionString)
            Try
                connection.Open()
                Dim command As New OleDbCommand("SELECT * FROM TBLNonReportableItems", connection)
                Dim adapter As New OleDbDataAdapter(command)
                Dim dataTable As New DataTable()
                adapter.Fill(dataTable)
                Return dataTable
            Catch ex As Exception
                MessageBox.Show($"Error retrieving data from TBLNonReportableItems: {ex.Message}")
                Return Nothing
            Finally
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                End If
            End Try
        End Using
    End Function

    ' Function to get data from TBLLocationCustodian table
    Public Function GetLocationCustodianData() As DataTable
        Using connection As OleDbConnection = OpenConnection()
            If connection Is Nothing Then
                MessageBox.Show("Unable to connect to the database.")
                Return Nothing
            End If

            Dim command As New OleDbCommand("SELECT * FROM TBLLocationCustodian", connection)
            Dim adapter As New OleDbDataAdapter(command)
            Dim dataTable As New DataTable()

            Try
                adapter.Fill(dataTable)
                Return dataTable
            Catch ex As Exception
                MessageBox.Show($"Error retrieving data from TBLLocationCustodian: {ex.Message}")
                Return Nothing
            Finally
                connection.Close()
            End Try
        End Using
    End Function

    ' Function to get data from TBLItems table
    Public Function GetItemsData() As DataTable
        Using connection As OleDbConnection = OpenConnection()
            If connection Is Nothing Then
                MessageBox.Show("Unable to connect to the database.")
                Return Nothing
            End If

            Dim command As New OleDbCommand("SELECT * FROM TBLItems", connection)
            Dim adapter As New OleDbDataAdapter(command)
            Dim dataTable As New DataTable()

            Try
                adapter.Fill(dataTable)
                Return dataTable
            Catch ex As Exception
                MessageBox.Show($"Error retrieving data from TBLItems: {ex.Message}")
                Return Nothing
            Finally
                connection.Close()
            End Try
        End Using
    End Function
    Public Sub InsertNonReportableItem(itemDetails As Dictionary(Of String, Object))
        Using connection As OleDbConnection = OpenConnection()
            If connection Is Nothing Then Return

            ' Start building the SQL statement dynamically
            Dim fields As New List(Of String)
            Dim placeholders As New List(Of String)
            Dim parameters As New List(Of Object)

            ' For each possible field, check if it's in itemDetails and add it to the command
            For Each field As String In New String() {"Location", "ItemCode", "SerialNum", "MakeModel", "AcqDate", "Cost", "[count]", "GrandTotal"}
                If itemDetails.ContainsKey(field) Then
                    fields.Add(field)
                    placeholders.Add("?")
                    parameters.Add(itemDetails(field))
                End If
            Next

            Dim cmdText As String = $"INSERT INTO TBLNonReportableItems ({String.Join(", ", fields)}) VALUES ({String.Join(", ", placeholders)})"
            Using cmd As New OleDbCommand(cmdText, connection)
                ' Add parameters to the command
                For Each param As Object In parameters
                    cmd.Parameters.AddWithValue("?", param)
                Next

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub UpdateNonReportableItem(itemId As Integer, itemDetails As Dictionary(Of String, Object))
        ' Ensure the connection is opened
        Using connection As OleDbConnection = OpenConnection()
            If connection Is Nothing Then Return

            ' Initialize the list for SET clause components and parameters
            Dim setClauses As New List(Of String)
            Dim parameters As New List(Of Object)

            ' Dynamically construct the SET clause based on itemDetails
            For Each kvp As KeyValuePair(Of String, Object) In itemDetails
                ' Exclude the ID from the SET clause
                If Not kvp.Key.Equals("ID", StringComparison.OrdinalIgnoreCase) Then
                    setClauses.Add($"[{kvp.Key}] = ?") ' Use square brackets around column names
                    parameters.Add(kvp.Value)
                End If
            Next

            ' Combine the SET clauses into a single string
            Dim setClauseString As String = String.Join(", ", setClauses)

            ' Assuming this is the correct dynamic SQL generation
            Dim cmdText As String = $"UPDATE TBLNonReportableItems SET {String.Join(", ", setClauses)} WHERE ID = ?"

            Using cmd As New OleDbCommand(cmdText, connection)
                For Each value In parameters ' Assuming 'parameters' is a List(Of Object) corresponding to values
                    cmd.Parameters.AddWithValue("?", value)
                Next
                ' The last parameter is for the WHERE clause, matching the ID
                cmd.Parameters.AddWithValue("?", itemId)

                Debug.WriteLine("Executing command: " & cmd.CommandText)
                For i As Integer = 0 To cmd.Parameters.Count - 1
                    Debug.WriteLine($"Param {i} ({cmd.Parameters(i).ParameterName}): {cmd.Parameters(i).Value} [{cmd.Parameters(i).Value.GetType()}]")
                Next

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub




    Public Sub DeleteNonReportableItem(itemId As Integer)
        Using connection As OleDbConnection = OpenConnection()
            If connection Is Nothing Then Return

            Dim cmdText As String = "DELETE FROM TBLNonReportableItems WHERE ID = ?"
            Using cmd As New OleDbCommand(cmdText, connection)
                cmd.Parameters.AddWithValue("@ID", itemId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

End Module
