VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scrolling Credits"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2445
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   2445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Show Splash using RES file"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show About using Text file"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmAboutCredits.ShowAboutCredits Me
    
End Sub

Private Sub Command2_Click()
    
    frmAboutCredits.ShowAboutSplash "TEXT", 101
    
End Sub
