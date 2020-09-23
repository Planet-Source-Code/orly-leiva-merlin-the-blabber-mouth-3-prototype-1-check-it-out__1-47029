VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merlin The Blabber-Mouth 3"
   ClientHeight    =   6615
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Load File So Merlin Can Read"
      Height          =   615
      Left            =   6240
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6120
      Top             =   5520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6120
      Top             =   4080
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Erase The List"
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Make Merlin Do It!"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add the Stuff so Merlin Can do it"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   4080
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   600
      TabIndex        =   2
      Top             =   4560
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1560
      List            =   "Form1.frx":00DF
      TabIndex        =   7
      Text            =   "Make Merlin Do Some Stuff"
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   495
      Left            =   6240
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop Talking Merlin!"
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3495
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6165
      _Version        =   393217
      TextRTF         =   $"Form1.frx":040F
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Make Merlin Say it!"
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Press the Delete Button to Erase an item"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Type whatever you want Merlin to say"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   5280
      Top             =   3960
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu OADSMCR 
         Caption         =   "Open A Document So Merlin Can Read"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
frmAbout.Show
End Sub

Private Sub BAO_Click()
Agent1.Characters("Merlin").Balloon.Visible = True
BAO.Enabled = False
BAOF.Enabled = True
End Sub

Private Sub BAOF_Click()
Agent1.Characters("Merlin").Balloon.Visible = False
BAO.Enabled = True
BAOF.Enabled = False
End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Anim = "Merlin"
    Set Char = Agent1.Characters(Anim)
    Char.AutoPopupMenu = False
    Char.Show
    Char.Speak Text1.Text
End Sub


Private Sub Command3_Click()
    End
End Sub

Private Sub Command4_Click()
Text1.Text = ""
End Sub

Private Sub Command5_Click()
List1.AddItem Combo1.Text
End Sub

Private Sub Command6_Click()
If List1.ListCount = 0 Then Exit Sub
Agent1.Characters("Merlin").Show
Agent1.Characters("Merlin").Play List1.List(List1.ListIndex)
End Sub

Private Sub Command7_Click()
List1.Clear
End Sub

Private Sub Command8_Click()
CommonDialog1.ShowOpen

Text1.LoadFile CommonDialog1.FileName
End Sub

Private Sub Form_Load()
Anim = "Merlin"
    Agent1.Characters.Load Anim, Anim & ".acs"
    Set Char = Agent1.Characters(Anim)
    Char.AutoPopupMenu = False
    Char.Show
    Char.MoveTo 300, 300
    Char.Play "Greet"
    Char.Speak ("Hello")
    Char.Play "Explain"
    Char.Speak ("This is Merlin.  You can type whatever you want in the textbox and I will say it.  And pick anything in the combo box and Add it to the list and I'll do it.  But first you got to make sure that you click on the list and the item on the list you want me to do.  Otherwise the program might crash.  So have fun!")
    Command2.Visible = True
End Sub


Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If List1.ListCount = 0 Then Exit Sub
If KeyCode = vbKeyDelete Then
List1.RemoveItem List1.ListIndex
End If
End Sub


Private Sub OADSMCR_Click()
CommonDialog1.ShowOpen

Text1.LoadFile CommonDialog1.FileName
End Sub

Private Sub Timer1_Timer()
Agent1.Characters("Merlin").Speak ("OK.  I'll Shut Up.")
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Agent1.Characters("Merlin").Stop
Timer2.Enabled = False
End Sub
