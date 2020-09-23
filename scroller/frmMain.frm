VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "A Scroller"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "&Different Settings"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdScroller 
      Caption         =   "&Scroller !"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.PictureBox OUT 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   18.813
      ScaleMode       =   4  'Zeichen
      ScaleWidth      =   85.625
      TabIndex        =   0
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdScroller_Click()

    'Example Syntax

    'Scroller NameOfPictureBox, "Your Message", FontSize, Animationspeed, Oscillation, Should Colors constantly change ?, otherwise name the Textcolor here
    
    Scroller OUT, "If you like this little Scroller contact me: overkillpage@gmx.net", 12, 3, 0.7, True, vbBlue
        
    
End Sub

Private Sub Command1_Click()
    
    Scroller OUT, "Calling the Scroller Sub with different Settings", 25, 6, 4, False, vbRed

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    End
    
End Sub
