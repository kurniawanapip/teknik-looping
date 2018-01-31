VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form5 
   Caption         =   "Aplikasi Hitung Nilai"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form5"
   ScaleHeight     =   6090
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Output Data"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1935
      Left            =   600
      TabIndex        =   13
      Top             =   3720
      Width           =   5655
      Begin RichTextLib.RichTextBox rtbtampil 
         Height          =   1575
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2778
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         RightMargin     =   10000
         TextRTF         =   $"DAFTAR.frx":0000
      End
   End
   Begin VB.CommandButton cmdcetak 
      Caption         =   "Cetak"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdhitung 
      Caption         =   "Hitung"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Data"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   4335
      Begin VB.TextBox txtumum 
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
         Left            =   1800
         TabIndex        =   12
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtmatematik 
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
         Left            =   1800
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtpsikotes 
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
         Left            =   1800
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtnama 
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
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Nilai Umum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Nilai Matematik"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Nilai Psikotes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Peserta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "DATA NILAI PESERTA TES"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vnama(10), vket(10) As String
Dim vpsikotes(10), vumum(10), vmatematik(10), vrata(10) As Single
Dim i, j As Integer

Private Sub cmdcetak_Click()
Open "D:\daftartes.txt" For Output As #1
Print #1, "            DATA NILAI PESERTA TES"
Print #1, "======================================"
Print #1, " NO"; Tab(6); "NAMA"; Tab(25); "RATA2"; Tab(40); "KET"
Print #1, "======================================"

For i = 1 To j
  Print #1, i; Tab(8); vnama(i); Tab(30); vrata(i); Tab(50); vket(i)
Next i

Print #1, "======================================"

Close #1

rtbtampil.FileName = "D:\daftartes.txt"

End Sub

Private Sub cmdhitung_Click()
For i = 1 To j
vrata(i) = (vpsikotes(i) + vmatematik(i) + vumum(i)) / 3

If vrata(i) >= 75 Then
 vket(i) = "LULUS"
 Else
 vket(i) = "GAGAL"
End If

Next i

MsgBox "proses hitung selesai", vbOKOnly + vbInformation, "informasi"
End Sub

Private Sub cmdsimpan_Click()
j = j + 1
vnama(j) = txtnama.Text
vpsikotes(j) = Val(txtpsikotes.Text)
vumum(j) = Val(txtumum.Text)
vmatematik(j) = Val(txtmatematik.Text)

txtnama.Text = ""
txtpsikotes.Text = ""
txtumum.Text = ""
txtmatematik.Text = ""
txtnama.SetFocus
 MsgBox "Data berhasil di simpan.", vbInformation + vbOKOnly, "informasi"
End Sub

