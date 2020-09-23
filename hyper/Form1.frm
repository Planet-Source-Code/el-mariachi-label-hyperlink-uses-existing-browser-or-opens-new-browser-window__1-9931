VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vist My WebSite..."
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0000
   ScaleHeight     =   3390
   ScaleWidth      =   4785
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option2 
      Caption         =   "Use Existing Browser"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "New Browser"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Website"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Choose to open a new browser or use the existing browser window. Then just click on either the button or the hyperlink. Thats, it."
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "www.WarpEngine.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   530
      Left            =   360
      MouseIcon       =   "Form1.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim ws As String
    Dim opn As Boolean
    
    'grab the webiste name from the label
    ws = Label1.Caption
    
    'do i need to open a new browser or use existing one
    If Option1.Value = True Then
        opn = True
    Else
        opn = False
    End If
    
    Call OpenLink(ws, opn)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ws As String
    Dim opn As Boolean
    
    'grab the webiste name from the label
    ws = Label1.Caption
    
    'do i need to open a new browser or use existing one
    If Option1.Value = True Then
        opn = True
    Else
        opn = False
    End If
    
    If Button = vbLeftButton Then
        Label1.ForeColor = vbRed
        Call OpenLink(ws, opn)
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = vbBlue
End Sub
