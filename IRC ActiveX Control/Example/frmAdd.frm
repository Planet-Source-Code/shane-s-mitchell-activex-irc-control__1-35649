VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtEnd 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "*!*@*"
      Top             =   3480
      Width           =   5055
   End
   Begin VB.TextBox txtHost 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Text            =   "*"
      Top             =   2760
      Width           =   5055
   End
   Begin VB.TextBox txtIdent 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Text            =   "*"
      Top             =   2040
      Width           =   5055
   End
   Begin VB.TextBox txtNick 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Text            =   "*"
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "End Result:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Host"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Ident"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nick"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "The following information can be in wildcard form"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   3435
   End
   Begin VB.Label Label1 
      Caption         =   "Add new user to access list"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmMain.List1.AddItem txtNick & "!" & txtIdent & "@" & txtHost
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub txtHost_Change()
    RebuildFinal
End Sub

Private Sub txtIdent_Change()
    RebuildFinal
End Sub

Private Sub txtNick_Change()
    RebuildFinal
End Sub

Sub RebuildFinal()
    txtEnd = txtNick & "!" & txtIdent & "@" & txtHost
End Sub
