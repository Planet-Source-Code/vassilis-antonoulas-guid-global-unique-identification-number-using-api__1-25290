VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   345
      TabIndex        =   1
      Top             =   480
      Width           =   3555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get GUID"
      Height          =   360
      Left            =   2685
      TabIndex        =   0
      Top             =   990
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = GetGUID
End Sub
