VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AgentCtl.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmusername 
   Caption         =   "Enter your name."
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1088
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   608
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmusername.frx":0000
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   1440
      Top             =   480
   End
End
Attribute VB_Name = "frmusername"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Username input
'Written by UberJavaJacker and #uberGeek
'Special thanks to my good friend UberJavaJacker for contributing the username thing

Dim LoadRequest(2)
Dim Merlin As IAgentCtlCharacterEx

Private Sub Command1_Click()
'A little error handling
If RichTextBox1.Text = "" Then
Merlin.Speak "You must enter a name!"
Else
If RichTextBox1.Text > "" Then
'Saves the name into a .fba file
RichTextBox1.SaveFile "c:\username.fba", 1
Merlin.Speak "Thanks," & RichTextBox1.Text & ", your name has been saved"
MsgBox "Your name has been saved"
Unload Me
End If
End If
End Sub
'Standard load routine
Private Sub Form_Load()
Agent1.Characters.Load "Merlin", "merlin.acs"
Set Merlin = Agent1.Characters("Merlin")
Merlin.Speak "If you enter your name, I will, from now on, call you by it."
End Sub
