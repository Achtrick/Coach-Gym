Public Class Acceuil

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Ajouter_Membre.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        forfait_expirée.Show()
        Me.Hide()
    End Sub
End Class
