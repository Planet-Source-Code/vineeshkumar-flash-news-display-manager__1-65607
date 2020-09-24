VERSION 5.00
Begin VB.Form frmFeedback 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Feedback"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMailto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      MouseIcon       =   "frmcalc.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Text            =   "vineesh@macrosoftindia.com"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmcalc.frx":030A
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "Vineesh kumar NV"
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmcalc.frx":0614
         Height          =   1875
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hi all.....,"
         Height          =   255
         Left            =   540
         TabIndex        =   1
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3840
      Y1              =   3000
      Y2              =   3000
   End
End
Attribute VB_Name = "frmFeedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub txtMailto_Click()
Shell "explorer mailto:vineesh@macrosoftindia.com", vbNormalFocus
End Sub
