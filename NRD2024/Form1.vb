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

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub guna2DgvNonReportableItems_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles guna2DgvNonReportableItems.CellContentClick

    End Sub
End Class
