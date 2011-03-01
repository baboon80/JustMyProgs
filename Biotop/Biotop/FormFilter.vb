Public Class FormFilter

    Private propRes As String

    Public Property prop1() As String
        Get
            ' The Get property procedure is called when the value
            ' of a property is retrieved.
            Return propRes
        End Get
        Set(ByVal value As String)
            ' The Set property procedure is called when the value 
            ' of a property is modified.  The value to be assigned
            ' is passed in the argument to Set.
            propRes = value
        End Set
    End Property



    Private Sub FormFilter_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Button1.Focus()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        propRes = Me.TextBox1.Text

        Call Close()

    End Sub

    Private Sub FormFilter_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        If propRes <> "" Then
            Me.TextBox1.Text = propRes
        Else
            Me.TextBox1.Text = Now.ToString("yyyy-MM-dd")
        End If
        RadioButton1.Checked = True
        Me.TextBox1.Focus()
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged

        If Me.TextBox1.Text = "" Then
            Me.TextBox1.Text = Now.ToString("yyyy-MM-dd")
        End If
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Return Then
            propRes = Me.TextBox1.Text

            Call Close()
        End If
    End Sub
End Class