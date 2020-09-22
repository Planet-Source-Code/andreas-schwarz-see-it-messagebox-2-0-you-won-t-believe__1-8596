VERSION 5.00
Begin VB.Form Msg 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Msg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdButton 
      Caption         =   "$"
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "$"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "$"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   1
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture2 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   2
      Top             =   0
      Width           =   0
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   3
      Left            =   120
      Picture         =   "Msg.frx":2CFA
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   2
      Left            =   120
      Picture         =   "Msg.frx":35C4
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   1
      Left            =   120
      Picture         =   "Msg.frx":3E8E
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "Msg.frx":4758
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lbText 
      BackStyle       =   0  'Transparent
      Caption         =   $"Msg.frx":5022
      Height          =   1335
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Msg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ButtonResult As MsgResult


Public Enum MsgType
    MsgError = 16
    MsgAsk = 32
    MsgWarning = 48
    MsgInfo = 64
    
    MsgOkOnly = 0
    MsgOkCancel = 1
    MsgOkIgnore = 2
    MsgYesNo = 3
    MsgYesNoCancel = 4
End Enum
    
Public Enum MsgResult
    MsgOk = 1
    MsgCancel = 2
    MsgIgnore = 3
    MsgYes = 4
    MsgNo = 5
End Enum


Function dwMessageBox(Prompt As String, Optional Flags As MsgType = 0, Optional PromptCaption As String, Optional iForm As Form) As MsgResult
    
    'Display Icon
    If Flags < 16 Then 'NoIcon
            
    ElseIf Flags < 32 Then 'Error-Icon
    
        IconImage(0).Visible = True
        Flags = Flags - 16
        
    ElseIf Flags < 48 Then 'Ask-Icon
        
        IconImage(1).Visible = True
        Flags = Flags - 32
        
    ElseIf Flags < 64 Then 'Warning-Icon
    
        IconImage(2).Visible = True
        Flags = Flags - 48
        
    ElseIf Flags < 80 Then 'Info-Icon
    
        IconImage(3).Visible = True
        Flags = Flags - 64
        
    End If
    
    'Displaytext
    
    lbText.Caption = Prompt
    Caption = PromptCaption
           
    'Initialize Buttons
    
    Select Case Flags
        Case 0
            SetButton 0, "OK"
        Case 1
            SetButton 0, "OK"
            SetButton 1, "Abbruch"
        Case 2
            SetButton 0, "OK"
            SetButton 1, "Ignorieren"
        Case 3
            SetButton 0, "Ja"
            SetButton 1, "Nein"
        Case 4
            SetButton 0, "Ja"
            SetButton 1, "Nein"
            SetButton 2, "Abbruch"
    End Select

   
    Msg.Show 1, iForm
    
    dwMessageBox = ButtonResult
    
    Unload Me
   
End Function
Private Sub SetButton(ButtonIndex As Integer, PromptText As String)
    cmdButton(ButtonIndex).Caption = PromptText
    cmdButton(ButtonIndex).Visible = True
End Sub

Private Sub cmdButton_Click(Index As Integer)
Select Case cmdButton(Index).Caption
    Case "Ok"
        ButtonResult = MsgOk
    Case "Abbruch"
        ButtonResult = MsgCancel
    Case "Ignorieren"
        ButtonResult = MsgIgnore
    Case "Ja"
        ButtonResult = MsgYes
    Case "Nein"
        ButtonResult = MsgNo
End Select
Me.Hide

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
               
End Sub

Private Sub Form_Load()
    For t = 0 To cmdButton.Count - 1
        cmdButton(t).Visible = False
    Next t
    For t = 0 To IconImage.Count - 1
        IconImage(t).Visible = False
    Next t
        
        
        
End Sub

