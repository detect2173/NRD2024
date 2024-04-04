Imports System.Data.OleDb
Imports System.Globalization

Public Class Form1

    Private displayingSearchResults As Boolean = False
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Apply the gradient background to th form
        GradientBackground.ApplyGradient(Me, "#C5ADC5", "#B2B5E0")
        ' Add any initialization after the InitializeComponent() call.
        AddHandler Guna2Panel1.Paint, AddressOf Guna2Panel1_Paint
    End Sub


    Private Sub Guna2Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Guna2Panel1.Paint
        Dim semiTransparentColor As Color = Color.FromArgb(64, Color.Gray)
        Dim semiTransparentPen As New Pen(semiTransparentColor, 2)
        e.Graphics.DrawRectangle(semiTransparentPen, 1, 1, Guna2Panel1.Width - 2, Guna2Panel1.Height - 2)
        semiTransparentPen.Dispose()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        guna2DgvNonReportableItems.DataSource = GetNonReportableItems()
        ' Populate ComboBox with table names
        cmbTableSelection.Items.Add("NonReportableItems")
        cmbTableSelection.Items.Add("Locations")
        cmbTableSelection.Items.Add("Items")
    End Sub

    Private Sub TabPage1_Enter(sender As Object, e As EventArgs) Handles TabPage1.Enter
        If Not displayingSearchResults Then
            guna2DgvNonReportableItems.DataSource = GetNonReportableItems()
            AdjustDgvForNonReportableItems()
        End If
    End Sub

    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        If Not displayingSearchResults Then
            guna2DgvLocationCustodian.DataSource = GetLocationCustodianData()
            AdjustDgvForLocationCustodianAndItems()
        End If
    End Sub

    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        If Not displayingSearchResults Then
            guna2DgvItems.DataSource = GetItemsData()
            AdjustDgvForLocationCustodianAndItems()
        End If
    End Sub


    Private Sub AdjustDgvForNonReportableItems()
        ' Example for guna2DgvNonReportableItems where autosize is set to column header
        For Each column As DataGridViewColumn In guna2DgvNonReportableItems.Columns
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
        Next
    End Sub

    Private Sub AdjustDgvForLocationCustodianAndItems()
        ' Assuming guna2DgvLocationCustodian and guna2DgvItems need the last column to stretch
        Dim dgvList As New List(Of Guna.UI2.WinForms.Guna2DataGridView) From {guna2DgvLocationCustodian, guna2DgvItems}

        For Each dgv As Guna.UI2.WinForms.Guna2DataGridView In dgvList
            If dgv.Columns.Count > 0 Then
                ' Set all but the last column to autosize based on column header
                For i As Integer = 0 To dgv.Columns.Count - 2
                    dgv.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                Next

                ' Set the last column to fill
                With dgv.Columns(dgv.Columns.Count - 1)
                    .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                End With
            End If
        Next
    End Sub
    Private Sub dgvNonReportableItems_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles guna2DgvNonReportableItems.CellClick
        ' Check if the click is on a row, not the column header
        If e.RowIndex >= 0 Then
            ' Get the selected row
            Dim selectedRow As DataGridViewRow = guna2DgvNonReportableItems.Rows(e.RowIndex)

            ' Populate the textboxes
            txtID.Text = selectedRow.Cells("ID").Value.ToString()
            txtLocation.Text = If(selectedRow.Cells("Location").Value IsNot Nothing, selectedRow.Cells("Location").Value.ToString(), String.Empty)
            txtItemCode.Text = If(selectedRow.Cells("ItemCode").Value IsNot Nothing, selectedRow.Cells("ItemCode").Value.ToString(), String.Empty)
            txtSerialNum.Text = If(selectedRow.Cells("SerialNumber").Value IsNot Nothing, selectedRow.Cells("SerialNumber").Value.ToString(), String.Empty)
            txtMakeModel.Text = If(selectedRow.Cells("MakeModel").Value IsNot Nothing, selectedRow.Cells("MakeModel").Value.ToString(), String.Empty)
            If Not IsDBNull(selectedRow.Cells("AcqDate").Value) Then
                ' Assuming 'AcqDate' is formatted as a date in the DataGridView
                txtAcqDate.Text = Convert.ToDateTime(selectedRow.Cells("AcqDate").Value).ToString("MM/dd/yyyy")
            Else
                txtAcqDate.Text = String.Empty
            End If
            If Not IsDBNull(selectedRow.Cells("Cost").Value) Then
                ' Format as currency
                txtCost.Text = Convert.ToDecimal(selectedRow.Cells("Cost").Value).ToString("C2")
            Else
                txtCost.Text = String.Empty ' or "$0.00" if you prefer to show zero as the default
            End If
            txtCount.Text = If(selectedRow.Cells("count").Value IsNot Nothing, selectedRow.Cells("count").Value.ToString(), String.Empty)
            If Not IsDBNull(selectedRow.Cells("GrandTotal").Value) Then
                ' Format as currency
                txtGrandTotal.Text = Convert.ToDecimal(selectedRow.Cells("GrandTotal").Value).ToString("C2")
            Else
                txtGrandTotal.Text = String.Empty ' or "$0.00" if you prefer to show zero as the default
            End If

            ' Since txtID is read-only, there's no need to set it again if you don't want to
        End If
    End Sub

    Private Sub btnNewNRI_Click(sender As Object, e As EventArgs) Handles btnNewNRI.Click
        ClearNRIFORM()
        ' Optional: Focus the first field after clearing
        txtLocation.Focus()
    End Sub

    Private Sub ClearNRIFORM()
        txtID.Clear()
        txtLocation.Clear()
        txtItemCode.Clear()
        txtSerialNum.Clear()
        txtMakeModel.Clear()
        txtAcqDate.Clear()
        txtCost.Clear()
        txtCount.Clear()
        txtGrandTotal.Clear()
    End Sub

    Private Sub btnSaveNRI_Click(sender As Object, e As EventArgs) Handles btnSaveNRI.Click
        ' Prepare variables to store parsed numeric values
        Dim costValue As Decimal = 0
        Dim grandTotalValue As Decimal = 0
        Dim countValue As Integer = 0

        ' Attempt to parse Cost and GrandTotal from currency format; Count as integer
        Dim parseCostSuccessful As Boolean = Decimal.TryParse(txtCost.Text, NumberStyles.Currency, CultureInfo.CurrentCulture, costValue)
        Dim parseGrandTotalSuccessful As Boolean = Decimal.TryParse(txtGrandTotal.Text, NumberStyles.Currency, CultureInfo.CurrentCulture, grandTotalValue)
        Dim parseCountSuccessful As Boolean = Integer.TryParse(txtCount.Text, countValue)

        ' Create the itemDetails dictionary with parsed data and checks for empty strings
        Dim itemDetails As New Dictionary(Of String, Object) From {
        {"Location", txtLocation.Text},
        {"ItemCode", txtItemCode.Text},
        {"SerialNum", txtSerialNum.Text},
        {"MakeModel", txtMakeModel.Text},
        {"AcqDate", If(String.IsNullOrWhiteSpace(txtAcqDate.Text), DBNull.Value, If(DateTime.TryParse(txtAcqDate.Text, New DateTime()), DateTime.Parse(txtAcqDate.Text), DBNull.Value))},
        {"Cost", If(parseCostSuccessful, costValue, DBNull.Value)},
        {"Count", If(parseCountSuccessful, countValue, DBNull.Value)},
        {"GrandTotal", If(parseGrandTotalSuccessful, grandTotalValue, DBNull.Value)}
    }

        ' Check if txtID is empty to determine insert or update
        If String.IsNullOrWhiteSpace(txtID.Text) Then
            ' Insert new record
            InsertNonReportableItem(itemDetails)
            MessageBox.Show("Record inserted successfully.")
        Else
            ' Update existing record
            Dim itemId As Integer = Integer.Parse(txtID.Text)
            UpdateNonReportableItem(itemId, itemDetails)
            MessageBox.Show("Record updated successfully.")
        End If

        ' Refresh DataGridView to reflect the changes
        RefreshDataGridView()
    End Sub

    Private Sub btnDeleteNRI_Click(sender As Object, e As EventArgs) Handles btnDeleteNRI.Click
        If guna2DgvNonReportableItems.CurrentRow IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(txtID.Text) Then
            Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this item?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If result = DialogResult.Yes Then
                Try
                    ' Attempt to delete the record
                    DeleteNonReportableItem(Convert.ToInt32(txtID.Text))
                    MessageBox.Show("Record deleted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ' Log the error or handle it as necessary
                Finally
                    RefreshDataGridView() ' Ensure the DataGridView is always refreshed
                    ClearNRIFORM() ' Clear the form
                End Try
            End If
        Else
            MessageBox.Show("Please select a record to delete.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub


    Private Sub RefreshDataGridView()
        guna2DgvNonReportableItems.DataSource = GetNonReportableItems()
    End Sub

    Function GetFieldNamesForTable(tableName As String) As List(Of String)
        Dim fieldNames As New List(Of String)

        ' Add the field names for each table
        Select Case tableName
            Case "TBLNonReportableItems"
                fieldNames.AddRange(New String() {"ID", "Location", "ItemCode", "SerialNumber", "MakeModel", "AcqDate", "Cost", "count", "GrandTotal"})
            Case "TBLLocationCustodian"
                fieldNames.AddRange(New String() {"ID", "custodianlname", "custodianfname", "location"})
            Case "TBLItems"
                fieldNames.AddRange(New String() {"ID", "ItemCode", "ItemDescription"})
        End Select

        Return fieldNames
    End Function

    Private Sub cmbTableSelection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTableSelection.SelectedIndexChanged
        ' Clear existing items and enable cmbFieldSelection
        cmbFieldSelection.Items.Clear()
        cmbFieldSelection.Enabled = True
        If cmbTableSelection.SelectedIndex >= 0 Then
            ' Determine the selected table
            Dim selectedTable As String = cmbTableSelection.SelectedItem.ToString()

            ' Map the user-friendly names to actual table names if necessary
            Dim tableNameMap As New Dictionary(Of String, String) From {
                {"NonReportableItems", "TBLNonReportableItems"},
                {"Locations", "TBLLocationCustodian"},
                {"Items", "TBLItems"}
            }

            Dim actualTableName As String = tableNameMap(selectedTable)

            ' Dynamically retrieve field names based on the selected table
            Dim fieldNames As List(Of String) = GetFieldNamesForTable(actualTableName)

            ' Debugging: output the count of field names retrieved
            Debug.WriteLine("Number of fields retrieved: " & fieldNames.Count)

            ' Add "All Fields" option
            cmbFieldSelection.Items.Add("All Fields")

            ' Populate cmbFieldSelection with actual field names
            For Each fieldName In fieldNames
                cmbFieldSelection.Items.Add(fieldName)
                ' Debugging: output each field name to the debug console
                Debug.WriteLine("Added field: " & fieldName)
            Next

            ' Select the "All Fields" option by default
            cmbFieldSelection.SelectedIndex = 0

            ' Open the relevant tab
            Select Case actualTableName
                Case "TBLNonReportableItems"
                    TabControl1.SelectedTab = TabPage1
                Case "TBLLocationCustodian"
                    TabControl1.SelectedTab = TabPage2
                Case "TBLItems"
                    TabControl1.SelectedTab = TabPage3
                Case Else
                    ' If no matching case is found, output a message to the debug console
                    Debug.WriteLine("No matching case for the selected table: " & actualTableName)
            End Select
        End If


    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        ' Get user selections
        Dim userSelectedTable As String = cmbTableSelection.SelectedItem.ToString()
        Dim selectedField As String = cmbFieldSelection.SelectedItem.ToString()
        Dim searchText As String = txtSearch.Text.Trim()

        ' Map user-friendly table name to actual database table name
        Dim actualTableName As String = MapUserSelectionToTableName(userSelectedTable)

        ' Determine the target DataGridView for displaying the search results
        Dim targetDGV As DataGridView = GetDataGridViewBasedOnTable(actualTableName)

        If targetDGV Is Nothing Then
            MessageBox.Show("No corresponding DataGridView found for the selected table.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Construct the search query based on user input
        Dim query As String
        If selectedField = "All Fields" Then
            query = ConstructSearchAllFieldsQuery(actualTableName, searchText)
        Else
            ' For specific field search, it's better to use parameterized queries, but this is a simplified version
            query = $"SELECT * FROM [{actualTableName}] WHERE [{selectedField}] LIKE '%{searchText}%'"
        End If
        displayingSearchResults = True
        ' Execute the search query and display the results in the appropriate DataGridView
        ExecuteSearchQuery(query, targetDGV)

    End Sub


    Private Function MapUserSelectionToTableName(userSelection As String) As String
        Select Case userSelection
            Case "NonReportableItems"
                Return "TBLNonReportableItems"
            Case "Locations"
                Return "TBLLocationCustodian" ' Make sure this is the correct table name
            Case "Items"
                Return "TBLItems" ' Make sure this is the correct table name
            Case Else
                Throw New Exception("Invalid selection.")
        End Select
    End Function

    Private Function ConstructSearchQuery(userSelection As String, searchField As String, searchText As String) As String
        Dim tableName = MapUserSelectionToTableName(userSelection)

        If searchField = "All Fields" Then
            ' Assuming ConstructSearchAllFieldsQuery is correctly implemented
            Return ConstructSearchAllFieldsQuery(tableName, searchText)
        Else
            ' Safe to assume parameters are used or input is sanitized to prevent SQL injection
            Return $"SELECT * FROM [{tableName}] WHERE [{searchField}] LIKE '%{searchText}%'"
        End If
    End Function



    Private Function ConstructSearchAllFieldsQuery(tableName As String, searchText As String) As String
        Dim fields As List(Of String) = GetFieldNamesForTable(tableName)
        ' Remove the "ID" field or any field that should not be searched
        fields.Remove("ID")

        Dim searchClauses As New List(Of String)
        For Each field In fields
            searchClauses.Add($"[{field}] LIKE '%{searchText}%'")
        Next

        ' Join the individual search clauses with OR
        Dim searchQuery As String = String.Join(" OR ", searchClauses)
        Dim query As String = $"SELECT * FROM [{tableName}] WHERE {searchQuery}"

        Return query
    End Function
    Private Sub ExecuteSearchQuery(query As String, targetDataGridView As DataGridView)
        Dim connectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Application.StartupPath}\newNonReportableItems.accdb"
        Using conn As New OleDbConnection(connectionString)
            Try
                conn.Open()
                Using cmd As New OleDbCommand(query, conn)
                    ' Execute the query and load the result into a DataTable
                    Dim table As New DataTable()
                    Using reader As OleDbDataReader = cmd.ExecuteReader()
                        table.Load(reader)
                        ' Set the DataSource of the target DataGridView to display the results
                        targetDataGridView.DataSource = table
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show($"An error occurred while executing the search: {ex.Message}", "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                conn.Close()
            End Try
        End Using
    End Sub
    ' Corrected Dynamic DataGridView Selection Based on Table
    ' Ensure this method accurately returns the correct DataGridView for the given table name
    Private Function GetDataGridViewBasedOnTable(actualTableName As String) As DataGridView
        Select Case actualTableName
            Case "TBLNonReportableItems"
                Return guna2DgvNonReportableItems
            Case "TBLLocationCustodian"
                Return guna2DgvLocationCustodian
            Case "TBLItems"
                Return guna2DgvItems
            Case Else
                Return Nothing
        End Select
    End Function



    Private Sub resetSearch()
        cmbTableSelection.SelectedIndex = 0
        cmbFieldSelection.SelectedIndex = -1
        txtSearch.ResetText()
        cmbTableSelection.SelectedIndex = -1

    End Sub

    Private Sub guna2DgvLocationCustodian_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles guna2DgvLocationCustodian.CellClick
        ' Check if the click is on a row, not the column header
        If e.RowIndex >= 0 Then
            ' Get the selected row
            Dim selectedRow As DataGridViewRow = guna2DgvLocationCustodian.Rows(e.RowIndex)

            ' Populate the textboxes
            txtIDLocations.Text = selectedRow.Cells("ID").Value.ToString()
            txtLocationLocations.Text = If(selectedRow.Cells("location").Value IsNot Nothing, selectedRow.Cells("location").Value.ToString(), String.Empty)
            txtLNameLocations.Text = If(selectedRow.Cells("custodianLname").Value IsNot Nothing, selectedRow.Cells("custodianLname").Value.ToString(), String.Empty)
            txtFNameLocations.Text = If(selectedRow.Cells("custodianFname").Value IsNot Nothing, selectedRow.Cells("custodianFname").Value.ToString(), String.Empty)
        End If
    End Sub

    Private Sub btnresetSearch_Click(sender As Object, e As EventArgs) Handles btnresetSearch.Click
        resetSearch()
        displayingSearchResults = False
    End Sub

    Private Sub btnNewLocations_Click(sender As Object, e As EventArgs) Handles btnNewLocations.Click
        ClearLocationsForm()
    End Sub



    Private Sub btnSaveLocations_Click(sender As Object, e As EventArgs) Handles btnSaveLocations.Click
        Dim locationData As New Dictionary(Of String, Object)
        locationData.Add("location", txtLocationLocations.Text)
        locationData.Add("custodianLname", txtLNameLocations.Text)
        locationData.Add("custodianFname", txtFNameLocations.Text)

        Try
            ' Assuming txtIDLocations contains the ID to update
            If String.IsNullOrWhiteSpace(txtIDLocations.Text) Then
                ' Insert new record
                InsertLocation(locationData)
                MessageBox.Show("Location inserted successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                ' Update existing record
                Dim id As Integer = Integer.Parse(txtIDLocations.Text)
                UpdateLocation(id, locationData)
                MessageBox.Show("Location updated successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Clear the controls
        ClearLocationsForm()

        ' Refresh the DataGridView
        RefreshLocationCustodianDGV()
    End Sub

    Private Sub ClearLocationsForm()
        ' Clear each TextBox in the Locations form.
        txtIDLocations.Clear()
        txtLocationLocations.Clear()
        txtLNameLocations.Clear()
        txtFNameLocations.Clear()
    End Sub

    Private Sub RefreshLocationCustodianDGV()
        ' Assuming you have a method to get the data and set it as the DataSource of the DataGridView
        guna2DgvLocationCustodian.DataSource = GetLocationCustodianData()
    End Sub



    Private Function collectLocationData() As Dictionary(Of String, Object)
        Dim locationData As New Dictionary(Of String, Object)
        ' Populate locationData with values from your form controls
        locationData("location") = txtLocationLocations.Text
        locationData("custodianLname") = txtLNameLocations.Text
        locationData("custodianFname") = txtFNameLocations.Text
        ' Add more fields here as necessary.
        Return locationData
    End Function
    Private Sub btnDeleteLocations_Click(sender As Object, e As EventArgs) Handles btnDeleteLocations.Click
        If Not String.IsNullOrWhiteSpace(txtIDLocations.Text) Then
            ' Confirm deletion
            Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this location?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If result = DialogResult.Yes Then
                Try
                    DeleteLocation(txtIDLocations.Text)
                    MessageBox.Show("Location deleted successfully.", "Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ' Optionally, refresh your data grid view or clear the form
                    ClearLocationsForm()
                    RefreshLocationCustodianDGV()
                Catch ex As Exception
                    MessageBox.Show("An error occurred while deleting the location: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select a location to delete.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Public Sub InsertLocation(locationData As Dictionary(Of String, Object))
        Using connection As New OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Application.StartupPath}\newNonReportableItems.accdb")
            connection.Open()
            Dim command As New OleDbCommand("INSERT INTO TBLLocationCustodian (custodianLname, custodianFname, location) VALUES (?, ?, ?)", connection)

            command.Parameters.AddWithValue("?", locationData("custodianLname"))
            command.Parameters.AddWithValue("?", locationData("custodianFname"))
            command.Parameters.AddWithValue("?", locationData("location"))

            command.ExecuteNonQuery()
        End Using
    End Sub



    Public Sub UpdateLocation(id As Integer, locationData As Dictionary(Of String, Object))
        Using connection As New OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Application.StartupPath}\newNonReportableItems.accdb")
            connection.Open()
            Dim command As New OleDbCommand("UPDATE TBLLocationCustodian SET custodianLname = ?, custodianFname = ?, location = ? WHERE ID = ?", connection)

            command.Parameters.AddWithValue("?", locationData("custodianLname"))
            command.Parameters.AddWithValue("?", locationData("custodianFname"))
            command.Parameters.AddWithValue("?", locationData("location"))
            command.Parameters.AddWithValue("?", id)

            command.ExecuteNonQuery()
        End Using
    End Sub



    Public Sub DeleteLocation(id As Integer)
        Using connection As New OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Application.StartupPath}\newNonReportableItems.accdb")
            connection.Open()
            Dim command As New OleDbCommand("DELETE FROM TBLLocationCustodian WHERE ID = ?", connection)

            command.Parameters.AddWithValue("?", id)

            command.ExecuteNonQuery()
        End Using
        guna2DgvLocationCustodian.DataSource = GetLocationCustodianData()
    End Sub



End Class
