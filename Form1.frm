VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.1#0"; "vbskpro2.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TEXT TO SPEECH "
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin vbskpro.Skinner Skinner1 
      Left            =   0
      Top             =   1920
      _ExtentX        =   1270
      _ExtentY        =   1270
      SysDisableSkinCaption=   "&Disable Skin"
   End
   Begin glxpbuttonz.UserButtonz UserButtonz1 
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LISTEN"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   65280
      ColorButtonUp   =   65280
      ColorButtonDown =   65280
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Set ObjTextToSpeech = CreateObject("SAPI.spVoice")
ObjTextToSpeech.speak Text1.Text


End Sub

Private Sub Timer1_Timer()
Picture1.Visible = True
Timer2.Enabled = True
Timer1.Enabled = True

End Sub

Private Sub Timer2_Timer()
Picture1.Visible = False
Picture2.Visible = True
Timer3.Enabled = True
Timer2.Enabled = False

End Sub

Private Sub Timer3_Timer()
Picture2.Visible = False
Picture3.Visible = True
Timer4.Enabled = True
Timer3.Enabled = True
End Sub

Private Sub Timer4_Timer()
Picture3.Visible = False
Picture4.Visible = True
Timer5.Enabled = True
Timer4.Enabled = False

End Sub

Private Sub Timer5_Timer()
Picture4.Visible = False
Picture5.Visible = True
Timer1.Enabled = True
Timer5.Enabled = False

End Sub

Private Sub Grid1_GotFocus()

End Sub

Private Sub UserButtonz1_Click()
Set ObjTextToSpeech = CreateObject("SAPI.spVoice")
ObjTextToSpeech.speak Text1.Text
End Sub
