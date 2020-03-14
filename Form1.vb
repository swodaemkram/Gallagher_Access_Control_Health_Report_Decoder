Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.IO

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)



    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

        OpenFileDialog1.Title = "Open Gallagher Test File"
        OpenFileDialog1.Filter = "json file |*.json"

        OpenFileDialog1.ShowDialog()

        Dim path As String = OpenFileDialog1.FileName
        TextBox1.Text = File.ReadAllText(path)
        Dim TestResult1_Name As String
        Dim TestResult1_Description As String
        Dim TestResult1_Severity As String
        Dim json As JObject = JObject.Parse(Me.TextBox1.Text)
        Dim TestResult1_FlaggedItemCount As String
        Dim TestResult1_Recommendations As String



        For q = 0 To 48

            ' ****************************************************************Test One******************************************************************************
            TestResult1_Name = json.SelectToken("content.TestResults[" & q & "].Name")
            TestResult1_Description = json.SelectToken("content.TestResults[" & q & "].Description")
            TestResult1_Severity = json.SelectToken("content.TestResults[" & q & "].Severity")
            TestResult1_FlaggedItemCount = json.SelectToken("content.TestResults[" & q & "].FlaggedItemCount")
            TestResult1_Recommendations = json.SelectToken("content.TestResults[" & q & "].Recommendations[0]")
            Me.TextBox2.Text = Me.TextBox2.Text & "Test" & q & vbCrLf
            Me.TextBox2.Text = Me.TextBox2.Text & TestResult1_Name & vbCrLf & TestResult1_Description & vbCrLf & TestResult1_Severity & vbCrLf & vbCrLf & TestResult1_FlaggedItemCount & " " & TestResult1_Recommendations
            Me.TextBox2.Text = TextBox2.Text & vbCrLf & "_____________________________________________________________________________________________________________" & vbCrLf
            '*****************************************************************End of Test One**********************************************************************
        Next


    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        PrintDialog1.ShowDialog()


    End Sub
End Class
Public Class JSON_result

End Class