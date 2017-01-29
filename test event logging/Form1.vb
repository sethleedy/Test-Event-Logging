Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ' Make sure the text boxes have something to compare
        If TextBox1.Text <> "" And TextBox2.Text <> "" Then
            ' Hard set username and password - Does it match ?
            If TextBox1.Text = "username" And TextBox2.Text = "password" Then
                ' clear the error label if it was used.
                Label1.Text = ""

                ' Login to whatever ....

                ' Check if login was good and
                ' Note Success to log
                writeToLog("User logged in", "information")

            Else
                ' Tell the user that the login was incorrect
                Label1.Text = "Error. Incorrect username/password."

                ' Write to log the infraction
                writeToLog("Incorrect login", "error")

            End If
        End If

    End Sub

    Private Sub LogEntriesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogEntriesToolStripMenuItem.Click

        Form2.Visible = True

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        Module1.deleteLog()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Setup the module1 code for use on first run of our code
        Module1.startUp()

        Timer1.Start()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Check for new data to display

        ' We can use the data brought over from module1 OR retrieve it from the Windows Events Log(to prove it was written).

        ' From Module1, display on Form2
        If Module1.newData <> "" Then
            Form2.TextBox1.Text += Module1.newData + vbCrLf ' Add to the textbox
            Module1.newData = "" ' empty the variable for the next timer tick

            ' The passed Module1.newData variable allows us to know when the data is different thus to update even Form3. Or we can compare against a Static variable
            ' From Events Log, display on Form3
            ' or we can just blindly update Form3.TextBox.Text on every tick, outside of this IF block, even when we do not need to.
            readFromLog()
        End If

        

    End Sub

    Private Sub LogEntriesByReadingLogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogEntriesByReadingLogToolStripMenuItem.Click

        Form3.Show()

    End Sub

    Private Sub ProgramToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProgramToolStripMenuItem.Click

        ' Open browser to view this program information
        'NavigateWebURL("https://docs.google.com/document/d/1JhziVUYnud1DB9R6O7ftUZuUgjPzb-6bXkmW7Q3X6j0/edit?usp=sharing", "default") ' Not Working
        Process.Start("https://docs.google.com/document/d/1JhziVUYnud1DB9R6O7ftUZuUgjPzb-6bXkmW7Q3X6j0/edit?usp=sharing")
    End Sub
End Class
