Option Explicit On
Option Strict On

Imports System.Data
Imports System.Data.OleDb


Module AccessDb_Connection
    Public Function GetConnectionString() As String
        Dim strCon As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath
        strCon &= "\CoachGym.accdb;Persist Security Info=False;"
        Return strCon
    End Function

    Public Con As New OleDbConnection(GetConnectionString())
    Public Cmd As OleDbCommand
    Public SQL As String = String.Empty

    Public Function PerformCRUD(Com As OleDbCommand) As DataTable
        Dim da As OleDbDataAdapter
        Dim dt As New DataTable()

        Try

            da = New OleDbDataAdapter
            da.SelectCommand = Com
            da.Fill(dt)

            Return dt

        Catch ex As Exception
            MessageBox.Show("Erreur:" & ex.Message, "Opération échouée",
                MessageBoxButtons.OK, MessageBoxIcon.Error)
            dt = Nothing
        End Try

        Return dt
    End Function

End Module
