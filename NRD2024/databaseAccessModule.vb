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
            For Each field As String In New String() {"Location", "ItemCode", "SerialNumber", "MakeModel", "AcqDate", "Cost", "[count]"}
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
        Using connection As OleDbConnection = OpenConnection()
            If connection Is Nothing Then
                Debug.WriteLine("Connection is Nothing.")
                Exit Sub
            End If

            ' Ensure connection is open
            If connection.State <> ConnectionState.Open Then
                Debug.WriteLine("Opening connection...")
                connection.Open()
            End If

            ' SQL command with named parameters
            Dim cmdText As String = "UPDATE TBLNonReportableItems SET " &
                                "[Location] = @Location, " &
                                "[ItemCode] = @ItemCode, " &
                                "[SerialNumber] = @SerialNumber, " &
                                "[MakeModel] = @MakeModel, " &
                                "[AcqDate] = @AcqDate, " &
                                "[Cost] = @Cost, " &
                                "[count] = @Count " &   ' "count" is a reserved word, so it is enclosed in brackets
                                "WHERE [ID] = @ID"

            Using cmd As New OleDbCommand(cmdText, connection)
                ' Add named parameters with values from the itemDetails dictionary
                cmd.Parameters.Add(New OleDbParameter("@Location", itemDetails("Location")))
                cmd.Parameters.Add(New OleDbParameter("@ItemCode", If(itemDetails.ContainsKey("ItemCode"), itemDetails("ItemCode"), DBNull.Value)))
                cmd.Parameters.Add(New OleDbParameter("@SerialNumber", If(itemDetails.ContainsKey("SerialNum"), itemDetails("SerialNum"), DBNull.Value)))
                cmd.Parameters.Add(New OleDbParameter("@MakeModel", If(itemDetails.ContainsKey("MakeModel"), itemDetails("MakeModel"), DBNull.Value)))
                cmd.Parameters.Add(New OleDbParameter("@AcqDate", If(itemDetails.ContainsKey("AcqDate"), itemDetails("AcqDate"), DBNull.Value)))
                cmd.Parameters.Add(New OleDbParameter("@Cost", If(itemDetails.ContainsKey("Cost"), itemDetails("Cost"), DBNull.Value)))
                cmd.Parameters.Add(New OleDbParameter("@Count", If(itemDetails.ContainsKey("Count"), itemDetails("Count"), DBNull.Value)))
                cmd.Parameters.Add(New OleDbParameter("@ID", itemId))

                ' Execute the update command
                Try
                    cmd.ExecuteNonQuery()
                    Debug.WriteLine("Command executed successfully.")
                Catch ex As OleDbException
                    ' Handle the exception (e.g., log the error, display a message, etc.)
                    Debug.WriteLine($"Error executing command: {ex.Message}")
                End Try
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
