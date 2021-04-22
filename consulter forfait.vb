Option Explicit On
Option Strict On
Imports System.Data.OleDb

Public Class Ajouter_Membre

    Private ID As String = ""
    Private forfait As String = ""

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        Acceuil.Show()
    End Sub

    Private Sub Ajouter_Membre_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ResetMe()
        LoadData()
    End Sub

    Private Sub ResetMe()
        Me.ID = ""
        nom.Text = ""
        numero.Text = ""
        musculation.Checked = False
        box.Checked = False
        cardio.Checked = False
        datefin.ResetText()
        forfait = ""
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

        Cmd.Parameters.AddWithValue("nom", nom.Text.Trim())
        Cmd.Parameters.AddWithValue("numero", numero.Text.Trim())
        Cmd.Parameters.AddWithValue("forfait", forfait)
        Cmd.Parameters.AddWithValue("datefin", datefin.Value)

        If str = "Update" And Not String.IsNullOrEmpty(Me.ID) Then
            Cmd.Parameters.AddWithValue("ID", Me.ID)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If String.IsNullOrEmpty(Me.nom.Text.Trim()) Or
        String.IsNullOrEmpty(Me.numero.Text.Trim()) Or
        (
        (musculation.Checked = False) And (box.Checked = False) And (cardio.Checked = False)
        ) Then
            MessageBox.Show("Vérifiez les champs obligatoire !", "Insertion de données",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If (musculation.Checked = True) Then
            forfait = forfait + musculation.Text + "\"
        End If

        If (box.Checked = True) Then
            forfait = forfait + box.Text + "\"
        End If

        If (cardio.Checked = True) Then
            forfait = forfait + cardio.Text
        End If

        SQL = "INSERT INTO Membre(Nom, Numero, Forfait, Date_fin) VALUES(@nom, @numero, @forfait, @datefin)"

        Execute(SQL, "Insert")

        MessageBox.Show("Données sauvegardées avec success.", "Insertion de données",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)

        LoadData()

        ResetMe()


    End Sub

    Private Sub LoadData(Optional keyword As String = "")

        SQL = "SELECT ID, Nom, Numero, Forfait, Date_fin FROM Membre " &
                "WHERE Nom LIKE @Keyword OR Numero LIKE @Keyword OR Forfait LIKE @Keyword ORDER BY ID ASC"

        Dim strKeyword As String = String.Format("%{0}%", keyword)

        Cmd = New OleDbCommand(SQL, Con)
        Cmd.Parameters.Clear()
        Cmd.Parameters.AddWithValue("keyword", strKeyword)

        Dim dt As DataTable = PerformCRUD(Cmd)

        With DataGridView1

            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .AutoGenerateColumns = True

            .DataSource = dt

            .Columns(0).HeaderText = "ID"
            .Columns(1).HeaderText = "Nom"
            .Columns(2).HeaderText = "Numero"
            .Columns(3).HeaderText = "Forfait"
            .Columns(4).HeaderText = "Datefin"

            .Columns(0).Width = 40
            .Columns(1).Width = 100
            .Columns(2).Width = 80
            .Columns(3).Width = 165
            .Columns(4).Width = 150

        End With



    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        Dim dgv As DataGridView = DataGridView1

        If e.RowIndex <> -1 Then

            Me.ID = Convert.ToString(dgv.CurrentRow.Cells(0).Value).Trim()
            nom.Text = Convert.ToString(dgv.CurrentRow.Cells(1).Value).Trim()
            numero.Text = Convert.ToString(dgv.CurrentRow.Cells(2).Value).Trim()
            Dim str As String = Convert.ToString(dgv.CurrentRow.Cells(3).Value).Trim()

            If str.Contains("Musculation") Then
                musculation.Checked = True
            Else
                musculation.Checked = False
            End If

            If str.Contains("Cardio") Then
                cardio.Checked = True
            Else
                cardio.Checked = False
            End If

            If str.Contains("Box") Then
                box.Checked = True
            Else
                box.Checked = False
            End If

            datefin.Value = Convert.ToDateTime(dgv.CurrentRow.Cells(4).Value)

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If DataGridView1.Rows.Count = 0 Then
            Exit Sub
        End If

        If String.IsNullOrEmpty(Me.ID) Then
            MessageBox.Show("Svp séléctionnez un membre !", "Modification de données",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If String.IsNullOrEmpty(Me.nom.Text.Trim()) Or
        String.IsNullOrEmpty(Me.numero.Text.Trim()) Or
        (
        (musculation.Checked = False) And (box.Checked = False) And (cardio.Checked = False)
        ) Then
            MessageBox.Show("Vérifiez les champs obligatoire !", "Modification de données",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If (musculation.Checked = True) Then
            forfait = forfait + musculation.Text + "\"
        End If

        If (box.Checked = True) Then
            forfait = forfait + box.Text + "\"
        End If

        If (cardio.Checked = True) Then
            forfait = forfait + cardio.Text
        End If

        SQL = "UPDATE Membre Set Nom = @nom, Numero = @numero, Forfait = @forfait, Date_fin = @datefin WHERE ID = @ID"

        Execute(SQL, "Update")

        MessageBox.Show("Données Modifiées avec success.", "Modification de données",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)

        LoadData()

        ResetMe()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If DataGridView1.Rows.Count = 0 Then
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

    Private Sub recherche_TextChanged_1(sender As Object, e As EventArgs) Handles recherche.TextChanged

        If Not String.IsNullOrEmpty(recherche.Text.Trim()) Then
            LoadData(Me.recherche.Text.Trim())
        Else
            LoadData()
        End If
        ResetMe()

    End Sub

    Private Sub DataGridView1_MouseHover(sender As Object, e As EventArgs) Handles DataGridView1.MouseHover
        ToolTip1.Show("Selectionnez un ligne pour Faire des modifications ou supprission.", DataGridView1)
    End Sub

End Class