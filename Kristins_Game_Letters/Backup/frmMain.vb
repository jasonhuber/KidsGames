
Public Class frmMain
    Inherits System.Windows.Forms.Form
    'http://www.ferngrovshs.qld.edu.au/events/shareit/downloads/Agent%20meets%20.NET.pdf
    Dim intCorrect, intWrong As Integer
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
    Friend WithEvents AxAgent1 As AxAgentObjects.AxAgent
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents lnklblOne As System.Windows.Forms.LinkLabel
    Friend WithEvents lnklblTwo As System.Windows.Forms.LinkLabel
    Friend WithEvents lnklblThree As System.Windows.Forms.LinkLabel
    Friend WithEvents lnklblFour As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel3 As System.Windows.Forms.LinkLabel
    Friend WithEvents AxTextToSpeech1 As AxHTTSLib.AxTextToSpeech
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.AxAgent1 = New AxAgentObjects.AxAgent
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.lnklblOne = New System.Windows.Forms.LinkLabel
        Me.lnklblTwo = New System.Windows.Forms.LinkLabel
        Me.lnklblThree = New System.Windows.Forms.LinkLabel
        Me.lnklblFour = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel3 = New System.Windows.Forms.LinkLabel
        Me.AxTextToSpeech1 = New AxHTTSLib.AxTextToSpeech
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        CType(Me.AxAgent1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxTextToSpeech1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AxAgent1
        '
        Me.AxAgent1.Enabled = True
        Me.AxAgent1.ImeMode = System.Windows.Forms.ImeMode.Off
        Me.AxAgent1.Location = New System.Drawing.Point(272, 408)
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
        'lnklblOne
        '
        Me.lnklblOne.Font = New System.Drawing.Font("Microsoft Sans Serif", 72.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lnklblOne.Location = New System.Drawing.Point(16, 8)
        Me.lnklblOne.Name = "lnklblOne"
        Me.lnklblOne.Size = New System.Drawing.Size(200, 136)
        Me.lnklblOne.TabIndex = 5
        Me.lnklblOne.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lnklblTwo
        '
        Me.lnklblTwo.Font = New System.Drawing.Font("Microsoft Sans Serif", 72.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lnklblTwo.Location = New System.Drawing.Point(288, 16)
        Me.lnklblTwo.Name = "lnklblTwo"
        Me.lnklblTwo.Size = New System.Drawing.Size(200, 136)
        Me.lnklblTwo.TabIndex = 6
        Me.lnklblTwo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lnklblThree
        '
        Me.lnklblThree.Font = New System.Drawing.Font("Microsoft Sans Serif", 72.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lnklblThree.Location = New System.Drawing.Point(56, 248)
        Me.lnklblThree.Name = "lnklblThree"
        Me.lnklblThree.Size = New System.Drawing.Size(200, 136)
        Me.lnklblThree.TabIndex = 7
        Me.lnklblThree.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lnklblFour
        '
        Me.lnklblFour.Font = New System.Drawing.Font("Microsoft Sans Serif", 72.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lnklblFour.Location = New System.Drawing.Point(344, 256)
        Me.lnklblFour.Name = "lnklblFour"
        Me.lnklblFour.Size = New System.Drawing.Size(200, 136)
        Me.lnklblFour.TabIndex = 8
        Me.lnklblFour.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 72.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel1.Location = New System.Drawing.Point(352, 256)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(200, 136)
        Me.LinkLabel1.TabIndex = 8
        Me.LinkLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 72.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel2.Location = New System.Drawing.Point(24, 8)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(200, 136)
        Me.LinkLabel2.TabIndex = 5
        Me.LinkLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LinkLabel3
        '
        Me.LinkLabel3.Font = New System.Drawing.Font("Microsoft Sans Serif", 72.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel3.Location = New System.Drawing.Point(64, 248)
        Me.LinkLabel3.Name = "LinkLabel3"
        Me.LinkLabel3.Size = New System.Drawing.Size(200, 136)
        Me.LinkLabel3.TabIndex = 7
        Me.LinkLabel3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'AxTextToSpeech1
        '
        Me.AxTextToSpeech1.Enabled = True
        Me.AxTextToSpeech1.Location = New System.Drawing.Point(184, 120)
        Me.AxTextToSpeech1.Name = "AxTextToSpeech1"
        Me.AxTextToSpeech1.OcxState = CType(resources.GetObject("AxTextToSpeech1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxTextToSpeech1.Size = New System.Drawing.Size(192, 192)
        Me.AxTextToSpeech1.TabIndex = 9
        Me.AxTextToSpeech1.Visible = False
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem3, Me.MenuItem4})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2})
        Me.MenuItem1.Text = "Options"
        '
        'MenuItem2
        '
        Me.MenuItem2.Checked = True
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "UseAgent"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "Correct: "
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 2
        Me.MenuItem4.Text = "Incorrect: "
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(568, 422)
        Me.Controls.Add(Me.AxTextToSpeech1)
        Me.Controls.Add(Me.lnklblFour)
        Me.Controls.Add(Me.lnklblThree)
        Me.Controls.Add(Me.lnklblTwo)
        Me.Controls.Add(Me.lnklblOne)
        Me.Controls.Add(Me.AxAgent1)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.LinkLabel2)
        Me.Controls.Add(Me.LinkLabel3)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmMain"
        Me.Text = "Kristin's Game"
        CType(Me.AxAgent1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxTextToSpeech1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "All of my Dims"
    Dim recSize As Rectangle
    Dim mainagent As AxAgentObjects.AxAgent
    Private mSpeaker As AgentObjects.IAgentCtlCharacter
    Dim gstrColor As String
    Dim gintWaiting As Integer
    Dim gintRandom As Integer
    Dim lnklblArray(3) As LinkLabel
    Dim gintBlinkCount As Integer
    Dim blDoneSpeakingandDoNewGame As Boolean
    Dim SpecialRequest As Object
#End Region

#Region "Setting up the form"
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MainMenu1.MenuItems(0).MenuItems(0).Checked = False
        Timer1.Enabled = False
        Call FullScreen(Me)
        Call SetupQuads()
        Try
            mainagent = AxAgent1
            mainagent.Characters.Load("Peedy", "C:\windows\msagent\chars\peedy.acs")
            mSpeaker = mainagent.Characters("Peedy")
            mSpeaker.SoundEffectsOn = True
        Catch mspeakerE As Exception
            MessageBox.Show("I am sorry that I could not load Peedy!. Please tell your mom or dad!", "Could not find Peedy!")
            End
        End Try

        Call GameStart()
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
        lnklblOne.Left = 0
        lnklblOne.Top = 0
        lnklblOne.Width = Me.Width / 2
        lnklblOne.Height = Me.Height / 2
        lnklblOne.BackColor = Color.Red

        lnklblTwo.Left = Me.Width / 2
        lnklblTwo.Top = 0
        lnklblTwo.Width = Me.Width / 2
        lnklblTwo.Height = Me.Height / 2
        lnklblTwo.BackColor = Color.White

        lnklblThree.Left = 0
        lnklblThree.Top = Me.Height / 2
        lnklblThree.Width = Me.Width / 2
        lnklblThree.Height = Me.Height / 2
        lnklblThree.BackColor = Color.Green

        lnklblFour.Left = Me.Width / 2
        lnklblFour.Top = Me.Height / 2
        lnklblFour.Width = Me.Width / 2
        lnklblFour.Height = Me.Height / 2
        lnklblFour.BackColor = Color.Yellow

    End Sub
    Private Sub PlaceLettersOnQuads()
        'right now we are only working on CAPITOL letters. This code should be modified to work on LOWER case letters!
        'Most of this code is to ensure that the letters arent repeated.
        Dim intRandom As Int16
        Dim i, j As Int16
        Dim blFoundLetterAlready As Boolean
        For i = 0 To 3
            Randomize(Now.Millisecond())
            intRandom = Math.Round(Rnd() * 25)
            blFoundLetterAlready = False
            If i <> 0 Then
                For j = 0 To i - 1
                    If Chr(65 + intRandom) = lnklblArray(j).Text Then
                        i = i - 1
                        blFoundLetterAlready = True
                        Exit For
                    End If
                Next
                If Not blFoundLetterAlready Then
                    lnklblArray(i).Text = Chr(65 + intRandom)
                End If
            Else
                lnklblArray(i).Text = Chr(65 + intRandom)
            End If
        Next
    End Sub
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        BlinkPanel(lnklblArray(gintRandom))
    End Sub
    Private Sub Clearallbuttherightone()
        Dim i As Int32
        For i = 0 To 3
            If lnklblArray(i).Text <> gstrColor Then
                lnklblArray(i).Text = ""
            End If
        Next
    End Sub
    Private Sub BlinkPanel(ByVal lnklblX As LinkLabel)
        If gintBlinkCount < 5 Then
            Timer2.Enabled = True
            Timer2.Interval = 1000
            gintBlinkCount += 1
            If lnklblX.Visible = True Then
                lnklblX.Visible = False
            Else
                lnklblX.Visible = True
            End If
        Else
            lnklblX.Visible = True
            Timer2.Enabled = False
        End If
    End Sub
#End Region

    Private Sub GameStart()
        lnklblArray(0) = lnklblOne
        lnklblArray(1) = lnklblTwo
        lnklblArray(2) = lnklblThree
        lnklblArray(3) = lnklblFour

        Call PlaceLettersOnQuads()
        Dim strColor As String
        Dim intRandom As Integer

        Randomize(Now.Millisecond())
        intRandom = Math.Round(Rnd() * 3)

        strColor = lnklblArray(intRandom).Text.ToString

        gstrColor = strColor
        gintRandom = intRandom
        If gstrColor = "A" Then
            gstrColor = "ay"
        End If
        Call SaySomething("Kristin please click on the letter " & gstrColor & "")
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        gintWaiting += 1
        If gintWaiting > 3 Then
            Call BlinkPanel(lnklblArray(gintRandom))
            Call SaySomething("The letter " & gstrColor & " is blinking at you!")
        End If
    End Sub
    Private Sub ClickAnylnllblLabel(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lnklblOne.Click, lnklblTwo.Click, lnklblThree.Click, lnklblFour.Click
        If lnklblArray(gintRandom) Is sender Then
            Timer1.Enabled = False
            Call Clearallbuttherightone()
            Call SaySomething("Good job Kristin, that is the letter " & gstrColor)
            SpecialRequest = mSpeaker.Speak("You got it right!")
            intCorrect += 1
            MainMenu1.MenuItems(1).Text = "Correct: " & intCorrect
        Else
            intWrong += 1
            MainMenu1.MenuItems(2).Text = "Incorrect: " & intWrong
            Call SaySomething("Good try, but we are looking for the letter " & gstrColor & "!")
        End If
    End Sub
    Private Sub AxAgent1_RequestComplete(ByVal sender As Object, ByVal e As AxAgentObjects._AgentEvents_RequestCompleteEvent) Handles AxAgent1.RequestComplete
        If e.request Is SpecialRequest Then
            'MsgBox(e.ToString)
            Call GameStart()
        End If
    End Sub
    Private Sub SaySomething(ByVal sText As String)
        If MainMenu1.MenuItems(0).MenuItems(0).Checked Then
            Try
                mSpeaker.Show(0)
                mSpeaker.Speak(sText)
                Timer1.Enabled = True
                gintWaiting = 1
            Catch e As Exception
                MessageBox.Show("I could not show the agent", "No agent to show")
            End Try
        Else
            AxTextToSpeech1.Speak(sText)
        End If
    End Sub
    Private Sub AxTextToSpeech1_SpeakingDone(ByVal sender As Object, ByVal e As System.EventArgs) Handles AxTextToSpeech1.SpeakingDone

    End Sub
End Class
