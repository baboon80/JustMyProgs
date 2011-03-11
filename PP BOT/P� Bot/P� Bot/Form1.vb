Public Class Form1

    Private Sub NeuePartieToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NeuePartieToolStripMenuItem.Click

    End Sub

    Private Sub PartieLadenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PartieLadenToolStripMenuItem.Click

    End Sub

    Private Sub BetVoyagerNonZeroToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BetVoyagerNonZeroToolStripMenuItem.Click

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private CurrentPosition As New System.Drawing.Point
    Private MouseButton As System.Windows.Forms.MouseButtons = Nothing

    Private Overloads Sub OnMouseDown(ByVal Sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown

        MyClass.MouseButton = e.Button()
        Me.Update()
        With MyClass.CurrentPosition
            .X = e.X()
            .Y = e.Y()
        End With

    End Sub

    Private Overloads Sub OnMouseMove(ByVal Sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove

        Select Case MouseButton
            Case Is = Windows.Forms.MouseButtons.Left
                MyClass.Top = Windows.Forms.Cursor.Position.Y() - MyClass.CurrentPosition.Y()
                MyClass.Left = Windows.Forms.Cursor.Position.X() - MyClass.CurrentPosition.X()
                Me.Update()
            Case Is = Nothing
                Me.Update()
                Exit Sub
        End Select
    End Sub

    Private Overloads Sub OnMouseUp(ByVal Sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
        MyClass.MouseButton = Nothing
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Close()
    End Sub

End Class
