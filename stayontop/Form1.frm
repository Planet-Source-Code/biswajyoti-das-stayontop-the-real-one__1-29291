VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   960
      Top             =   1.17920e5
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "askbiswa@hotmail.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1680
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4080
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":068A
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   2280
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":078C
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is by Biswajyoti Das
'If you use this code in your project then please let me know
'I am just kidding this code is nothin' .. I have just placed some stuff here and
'there. No need to mention my name .. but do send me feedback .. at my addy
'askbiswa@hotmail.com

'I developed this code because I have seen many people boasting that their code being
'the one that stays avove all of the other windows .. well I hae never found their
'claim to be true so I wrote this code .... in order to test my code just use any
'yes any stay on top form in the software world and I bet it won't fail ... what more
'it can be on top of even the menu items of the Windows start menu ... jus place the
'form around three inches above start menu and you see the menu items going behind it
'isn't this cool ...
'Just remember me Bis  and askbiswa@hotmail.com .... I know now my mailbox is going to
'be flooded with mails for more cool and simple stuff like this .... well I promise
'once I get started I know I will never stop ... so write back soon ... if you like
'it

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40


Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Form_Load()
 SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
Private Sub Label3_Click()
Shell ("start mailto:askbiswa@hotmail.com")
End Sub












































































































































































Private Sub Timer1_Timer()
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
