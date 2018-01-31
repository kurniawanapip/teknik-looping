VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "latihan looping"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "proses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Deret 10 pertama Bilangan genap"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vsuku, x As Single
Dim vhasil As String
Private Sub Command1_Click()
vsuku = 2
'x = 1

'Do While x <= 10
  'vhasil = vhasil + Str(vsuku)
  'vsuku = vsuku + 2
  'x = x + 1
  'Loop
 
 'Do Until x > 10
   'vhasil = vhasil + Str(vsuku)
   'vsuku = vsuku + 2
   'x = x + 1
   'Loop
   
   For x = 1 To 10
      vhasil = vhasil + Str(vsuku)
      vsuku = vsuku + 2
      
      Next x
      
  Text1.Text = vhasil
End Sub

