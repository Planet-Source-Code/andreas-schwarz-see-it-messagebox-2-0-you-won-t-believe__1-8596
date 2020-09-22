VERSION 5.00
Begin VB.Form Test 
   Caption         =   "MessageBox Version 2.0 Test"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Press"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim iResult As MsgResult
    
    If MessageBox("Do you want to start the new MessageBox Style ?", MsgAsk + MsgYesNo, "Welcome") = MsgYes Then
        MessageBox "You clicked 'YES'.", MsgInfo, "Other Caption"
    Else
        MessageBox "You clicked 'NO'.", MsgError
    End If
        
    

End Sub
