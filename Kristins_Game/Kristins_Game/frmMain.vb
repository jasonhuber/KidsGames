Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents pnlOne As System.Windows.Forms.Panel
    Friend WithEvents pnlTwo As System.Windows.Forms.Panel
    Friend WithEvents pnlThree As System.Windows.Forms.Panel
    Friend WithEvents pnlFour As System.Windows.Forms.Panel
    Friend WithEvents AxAgent1 As AxAgentObjects.AxAgent
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.pnlOne = New System.Windows.Forms.Panel()
        Me.pnlTwo = New System.Windows.Forms.Panel()
        Me.pnlThree = New System.Windows.Forms.Panel()
        Me.pnlFour = New System.Windows.Forms.Panel()
        Me.AxAgent1 = New AxAgentObjects.AxAgent()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        CType(Me.AxAgent1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlOne
        '
        Me.pnlOne.Location = New System.Drawing.Point(8, 8)
        Me.pnlOne.Name = "pnlOne"
        Me.pnlOne.Size = New System.Drawing.Size(272, 192)
        Me.pnlOne.TabIndex = 0
        '
        'pnlTwo
        '
        Me.pnlTwo.Location = New System.Drawing.Point(288, 8)
        Me.pnlTwo.Name = "pnlTwo"
        Me.pnlTwo.Size = New System.Drawing.Size(272, 192)
        Me.pnlTwo.TabIndex = 1
        '
        'pnlThree
        '
        Me.pnlThree.Location = New System.Drawing.Point(8, 208)
        Me.pnlThree.Name = "pnlThree"
        Me.pnlThree.Size = New System.Drawing.Size(272, 192)
        Me.pnlThree.TabIndex = 2
        '
        'pnlFour
        '
        Me.pnlFour.Location = New System.Drawing.Point(288, 208)
        Me.pnlFour.Name = "pnlFour"
        Me.pnlFour.Size = New System.Drawing.Size(272, 192)
        Me.pnlFour.TabIndex = 3
        '
        'AxAgent1
        '
        Me.AxAgent1.Enabled = True
        Me.AxAgent1.Location = New System.Drawing.Point(264, 400)
        Me.AxAgent1.Name = "AxAgent1"
        Me.AxAgent1.OcxState = CType(resources.GetObject("AxAgent1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxAgent1.Size = New System.Drawing.Size(32, 32)
        Me.AxAgent1.TabIndex = 4
        '
        'Timer1
        '
        Me.Timer1.Interval = 10000
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(568, 422)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.AxAgent1, Me.pnlFour, Me.pnlThree, Me.pnlTwo, Me.pnlOne})
        Me.Name = "frmMain"
        Me.Text = "Kristin's Game"
        CType(Me.AxAgent1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim recSize As Rectangle
    Dim mainagent As AxAgentObjects.AxAgent
    Private mSpeaker As AgentObjects.IAgentCtlCharacter
    Dim gstrColor As String
    Dim gintWaiting As Integer
    Dim gintRandom As Integer
    Dim pnlArray(3) As Panel
    Dim gintBlinkCount As Integer
    Dim gstrLastintRandom As Integer
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Enabled = False
        Dim largehand As New System.Windows.Forms.Cursor("largehand.cur")
        Me.Cursor = largehand
        Call FullScreen(Me)
        Call SetupQuads()
        Try
            mainagent = AxAgent1
            mainagent.Characters.Load("Peedy", "C:\WINDOWS\msagent\chars\peedy.acs")
            mSpeaker = mainagent.Characters("Peedy")
            mSpeaker.SoundEffectsOn = True

        Catch mspeakerE As Exception
            MessageBox.Show("I am sorry that I could not load Peedy!. Please tell your mom or dad!", "Could not find Peedy!")
            End
        End Try
        Call GameStart()
    End Sub

    Private Sub GameStart()
        pnlArray(0) = pnlOne
        pnlArray(1) = pnlTwo
        pnlArray(2) = pnlThree
        pnlArray(3) = pnlFour

        Dim strColor As String
        Dim intRandom As Integer
        Do
            Randomize(Now.Millisecond)
            intRandom = Rnd() * 3

            strColor = pnlArray(intRandom).BackColor.ToString
            strColor = Microsoft.VisualBasic.Mid(strColor, Microsoft.VisualBasic.InStr(strColor, "[") + 1)
            strColor = strColor.Substring(0, strColor.IndexOf("]"))

            gstrColor = strColor
            gintRandom = intRandom
        Loop Until gintRandom <> gstrLastintRandom
        gstrLastintRandom = gintRandom
        Try
            mSpeaker.Show(0)
            mSpeaker.Speak(gstrColor)
            mSpeaker.Speak("Kristin please show me the " & gstrColor & " square!")
            Select Case LCase(strColor)
                Case "green"
                    mSpeaker.Speak("Green is the color of Grass!")
                Case "red"
                    mSpeaker.Speak("Red is the color of tomatoes!")
                Case "yellow"
                    mSpeaker.Speak("Yellow is the color of the sun!")
                Case "blue"
                    mSpeaker.Speak("Blue is the color of water!")
            End Select
            Timer1.Enabled = True
            gintWaiting = 1
        Catch e As Exception
            MessageBox.Show("I could not show the agent", "No agent to show")
        End Try
    End Sub

    Private Sub FullScreen(ByVal x As Form)
        Dim recSize As Rectangle
        recSize = Screen.PrimaryScreen.WorkingArea
        x.Height = recSize.Height
        x.Width = recSize.Width
        x.Left = 0
        x.Top = 0
    End Sub
    Private Sub SetupQuads()

        Dim strColors(4) As System.Drawing.Color
        Dim intColor As Integer
        Dim tempColor As Integer
        strColors(0) = System.Drawing.Color.Red
        strColors(1) = System.Drawing.Color.Blue
        strColors(2) = System.Drawing.Color.Green
        strColors(3) = System.Drawing.Color.Yellow
        Randomize(Now.Millisecond)
        intColor = Rnd() * 3
        tempColor = intColor
        pnlOne.Left = 0
        pnlOne.Top = 0
        pnlOne.Width = Me.Width / 2
        pnlOne.Height = Me.Height / 2
        pnlOne.BackColor = strColors(intColor)


        Do Until strColors(intColor).ToString <> pnlOne.BackColor.ToString
            Randomize(Now.Millisecond)
            intColor = Rnd() * 3
        Loop

        pnlTwo.Left = Me.Width / 2
        pnlTwo.Top = 0
        pnlTwo.Width = Me.Width / 2
        pnlTwo.Height = Me.Height / 2
        pnlTwo.BackColor = strColors(intColor)

        Do Until strColors(intColor).ToString <> pnlOne.BackColor.ToString _
            And strColors(intColor).ToString <> pnlTwo.BackColor.ToString
            Randomize(Now.Millisecond)
            intColor = Rnd() * 3
        Loop

        pnlThree.Left = 0
        pnlThree.Top = Me.Height / 2
        pnlThree.Width = Me.Width / 2
        pnlThree.Height = Me.Height / 2
        pnlThree.BackColor = strColors(intColor)

        Do Until strColors(intColor).ToString <> pnlOne.BackColor.ToString _
            And strColors(intColor).ToString <> pnlTwo.BackColor.ToString _
            And strColors(intColor).ToString <> pnlThree.BackColor.ToString
            Randomize(Now.Millisecond)
            intColor = Rnd() * 3
        Loop

        pnlFour.Left = Me.Width / 2
        pnlFour.Top = Me.Height / 2
        pnlFour.Width = Me.Width / 2
        pnlFour.Height = Me.Height / 2
        pnlFour.BackColor = strColors(intColor)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        gintWaiting += 1
        If gintWaiting > 3 Then
            Try
                mSpeaker.Show(0)
                mSpeaker.Speak("OK we will try another one!")
                Call BlinkPanel(pnlArray(gintRandom))
                Call GameStart()
            Catch r As Exception
                MessageBox.Show("I could not load the speaker!", "Could not show agent!")
            End Try
        End If
        Timer1.Enabled = False
    End Sub
    Private Sub BlinkPanel(ByVal pnlX As Panel)
        If gintBlinkCount < 3 Then
            Timer2.Enabled = True
            gintBlinkCount += 1
            If pnlX.Visible = True Then
                pnlX.Visible = False
            Else
                pnlX.Visible = True
            End If
        End If

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        BlinkPanel(pnlArray(gintRandom))
        Timer2.Enabled = False
    End Sub

    Private Sub pnlOne_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlOne.Click, pnlTwo.Click, pnlThree.Click, pnlFour.Click
        If pnlArray(gintRandom) Is sender Then
            Timer1.Enabled = False
            Try
                mSpeaker.Show(0)
                mSpeaker.Speak("Good Job Kristin!")
                mSpeaker.Speak("Lets try another one!")
            Catch t As Exception
            End Try
            Call GameStart()
        Else
            Try
                mSpeaker.Show(0)
                mSpeaker.Speak("not that one!")
            Catch t As Exception
            End Try
        End If
    End Sub
End Class
