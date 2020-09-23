VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test colors"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "TEST colors"
      Height          =   675
      Left            =   570
      TabIndex        =   0
      Top             =   495
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ColorDlg Form1.hWnd, RGB(255, 255, 255), True
End Sub
