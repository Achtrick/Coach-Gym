Option Explicit On
Option Strict On
Imports System.Data.OleDb

Public Class forfait_expirée

    Private ID As String = ""

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        Acceuil.Show()
    End Sub

    Private Sub LoadData(Optional keyword As String = "")

        SQL = "SELECT ID, Nom, Numero, Forfait, Date_fin FROM Membre " &
                "WHERE Date_fin <= @Keyword ORDER BY Date_fin ASC"

        Cmd = New OleDbCommand(SQL, Con)
        Cmd.Parameters.Clear()
        Cmd.Parameters.AddWithValue("keyword", DateTime.Today.ToString)

        Dim dt As DataTable = PerformCRUD(Cmd)

        With DataGridView2

            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .AutoGenerateColumns = True

            .DataSource = dt

            .Columns(0).HeaderText = "ID"
            .Columns(1).HeaderText = "Nom"
            .Columns(2).HeaderText = "Numero"
            .Columns(3).HeaderText = "Forfait"
            .Columns(4).HeaderText = "Datefin"

            .Columns(0).Width = 200
            .Columns(1).Width = 200
            .Columns(2).Width = 200
            .Columns(3).Width = 200
            .Columns(4).Width = 200

        End With



    End Sub

    Private Sub forfait_expirée_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If DataGridView2.Rows.Count = 0 Then
            Exit Sub
        End If

        If String.IsNullOrEmpty(Me.ID) Then
            MessageBox.Show("Svp séléctionnez un membre !", "Suppression de données",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If MessageBox.Show("Est ce que vous êtes sur de supprimer ce membre ?", "Suppression de sonnées",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            SQL = "DELETE * FROM Membre WHERE ID = @ID"

            Execute(SQL, "Delete")

            MessageBox.Show("Données Modifiées avec success.", "Suppression de données",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)

            LoadData()

            ResetMe()

        End If
    End Sub

    Private Sub ResetMe()
        Me.ID = ""
    End Sub

    Private Sub Execute(MySQL As String, Optional Parameter As String = "")
        Cmd = New OleDbCommand(MySQL, Con)
        AddParameters(Parameter)
        PerformCRUD(Cmd)
    End Sub

    Private Sub AddParameters(str As String)
        Cmd.Parameters.Clear()

        If str = "Delete" And Not String.IsNullOrEmpty(Me.ID) Then
            Cmd.Parameters.AddWithValue("ID", Me.ID)
        End If

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim dgv As DataGridView = DataGridView2

        If e.RowIndex <> -1 Then

            Me.ID = Convert.ToString(dgv.CurrentRow.Cells(0).Value).Trim()

        End If
    End Sub
End Class