Imports System.Data.OleDb
Imports System.Globalization

Public Class Form1
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Apply the gradient background to the form
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
    End Sub

    Private Sub TabPage1_Enter(sender As Object, e As EventArgs) Handles TabPage1.Enter
        guna2DgvNonReportableItems.DataSource = GetNonReportableItems()
        AdjustDgvForNonReportableItems()
    End Sub

    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        guna2DgvLocationCustodian.DataSource = GetLocationCustodianData()
        AdjustDgvForLocationCustodianAndItems()
    End Sub

    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        guna2DgvItems.DataSource = GetItemsData()
        AdjustDgvForLocationCustodianAndItems()
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

End Class
