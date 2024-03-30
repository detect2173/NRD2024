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

End Module
