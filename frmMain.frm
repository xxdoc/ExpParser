VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Expression Parser"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Height          =   7935
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmMain.frx":0000
      Top             =   4800
      Width           =   20895
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmMain.frx":0006
      Top             =   2880
      Width           =   20895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   615
      Left            =   6600
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmMain.frx":000C
      Top             =   0
      Width           =   20895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim oExp As New CExpression
    oExp.Parse Me.Text1.Text
    If oExp.ErrDesc <> "" Then
        Me.Text2.Text = oExp.ErrDesc
    Else
        Me.Text2.Text = oExp.ToXML()
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then End
End Sub

Private Sub Form_Load()
    Me.Text1.Text = "[UPB($)]-   (_F([Bal])/""100""   + _C(""Name""))"
End Sub
