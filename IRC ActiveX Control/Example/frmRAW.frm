VERSION 5.00
Begin VB.Form frmRAW 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "RAW Data"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   5655
   End
   Begin VB.TextBox txtRaw 
      Appearance      =   0  'Flat
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmRAW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    On Error Resume Next
    txtRaw.Move 15, 15, Me.ScaleWidth - 30, Me.ScaleHeight - 15 - txtInput.Height
    txtInput.Move 15, Me.ScaleHeight - txtInput.Height - 15, Me.ScaleWidth - 30
    txtRaw.SelStart = Len(txtRaw.Text)
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmMain.irc.Raw txtInput.Text
        txtInput.Text = ""
    End If
End Sub
