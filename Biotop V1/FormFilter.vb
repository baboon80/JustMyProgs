Public Class FormFilter

    Private propDatum As String
    Private propInvert As Boolean

    Public Property Datum() As String
        Get
            ' The Get property procedure is called when the value
            ' of a property is retrieved.
            Return propDatum
        End Get
        Set(ByVal value As String)
            ' The Set property procedure is called when the value 
            ' of a property is modified.  The value to be assigned
            ' is passed in the argument to Set.
            propDatum = value
        End Set
    End Property

    Public Property Invert() As Boolean
        Get
            ' The Get property procedure is called when the value
            ' of a property is retrieved.
            Return propInvert
        End Get
        Set(ByVal value As Boolean)
            ' The Set property procedure is called when the value 
            ' of a property is modified.  The value to be assigned
            ' is passed in the argument to Set.
            propInvert = value
        End Set
    End Property

    Private Sub FormFilter_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Button1.Focus()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        CloseForm()
    End Sub

    Private Sub FormFilter_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        If propDatum <> "" Then
            Me.TextBox1.Text = propDatum
        Else
            'Me.TextBox1.Text = Now.ToString("yyyy-MM-dd")
        End If

        CheckBox1.Checked = propInvert
        Me.TextBox1.Focus()
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Return Then
            propDatum = Me.TextBox1.Text

            CloseForm()
        End If
    End Sub

    Private Sub CloseForm()
        propDatum = Me.TextBox1.Text
        propInvert = Me.CheckBox1.Checked
        Close()
    End Sub
End Class