VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Alert Box Demo"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Alert message"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "&Show alert"
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "This is a demo of the alert box."
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Alert(Text As String)
    ' Display a new alertbox
    Dim AlertBox As frmAlert
    Set AlertBox = New frmAlert
    AlertBox.DisplayAlert Text, 3000
    Me.SetFocus
End Sub

Private Sub Command1_Click()
    Alert Text1.Text
End Sub
