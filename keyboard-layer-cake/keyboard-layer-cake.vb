
Public Class frmKeyboardLayerCake
    Private WithEvents kbHook As New CakeKeyboardHook
    Private WithEvents oTimer As New Timer
    Private DieAfterMessage As Boolean = False

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.oTimer.Enabled = False
    End Sub
    Protected Overrides ReadOnly Property CreateParams() _
      As System.Windows.Forms.CreateParams
        Get
            Const WS_EX_NOACTIVATE As Integer = &H8000000
            Dim Result As CreateParams
            Result = MyBase.CreateParams
            Result.ExStyle = Result.ExStyle Or WS_EX_NOACTIVATE
            Return Result
        End Get
    End Property

    Private Sub kbHook_ShowToast(Message As String, ThenDie As Boolean) Handles kbHook.ShowToast
        Me.oTimer.Interval = 2000
        Me.Visible = True
        Me.Top = 0
        Me.Label1.AutoSize = True
        Me.TopMost = True
        Me.Label1.Text = Message
        'Me.Label2.Text = Message
        Me.Width = Me.Label1.Width
        Me.Height = Me.Label1.Height
        Me.oTimer.Enabled = True
        Me.oTimer.Start()
        DieAfterMessage = ThenDie
    End Sub

    Private Sub oTimer_Tick(sender As Object, e As EventArgs) Handles oTimer.Tick
        Me.Visible = False
        Me.oTimer.Enabled = False
        Me.oTimer.Stop()
        If DieAfterMessage Then
            End
        End If
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Visible = False

        If IO.File.Exists("keyboard-layer-cake.json") = False Then
            kbHook_ShowToast("keyboard-layer-cake.json not found.", True)
            Exit Sub
        End If
    End Sub
End Class

