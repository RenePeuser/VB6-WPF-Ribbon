VERSION 5.00
Object = "{7DC417DB-81C4-46F1-8B99-29EB92BF58D5}#1.0#0"; "WPFProgressbar.tlb"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8496
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8712
   LinkTopic       =   "Form1"
   ScaleHeight     =   8496
   ScaleWidth      =   8712
   StartUpPosition =   3  'Windows-Standard
   Begin WPFProgressbarCtl.WPFProgressbar WPFProgressbar1 
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1572
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      Minimum         =   "0"
      Maximum         =   "100"
      Value           =   "0"
      BackColor       =   "Control"
      ForeColor       =   "ControlText"
      Location        =   "10, 220"
      Name            =   "WPFProgressbar"
      Size            =   "131, 31"
      Object.TabIndex        =   "0"
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'Kein
      Height          =   5892
      Left            =   0
      ScaleHeight     =   5892
      ScaleWidth      =   9852
      TabIndex        =   1
      Top             =   3120
      Width           =   9852
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'Kein
      Height          =   2892
      Left            =   0
      ScaleHeight     =   2892
      ScaleWidth      =   9492
      TabIndex        =   0
      Top             =   -480
      Width           =   9492
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Sub Form_Load()
    
    
    Dim a As New RibbonControlForm.RibbonForm
        
    a.Show
    a.BorderStyle = 0
    Call SetParent(a.GetHandle, Me.Picture1.hWnd)
    a.Move 0, 0, Me.Picture1.Width, Me.Picture1.Height
    
    Dim b As New WPFtoVB6.InteropForm1
    
    b.Show
    b.BorderStyle = 0
    Call SetParent(b.GetHandle, Me.Picture2.hWnd)
    b.Move 0, -350, Me.Picture2.Width, Me.Picture2.Height

        
End Sub

Private Sub WPFButton1_MouseDown()

End Sub
