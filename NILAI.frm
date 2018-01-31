VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form4 
   Caption         =   "aplikasi hitung nilai semester"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form4"
   ScaleHeight     =   6960
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "output data"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2415
      Left            =   720
      TabIndex        =   13
      Top             =   4080
      Width           =   7095
      Begin RichTextLib.RichTextBox rtbtampil 
         Height          =   2055
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3625
         _Version        =   393217
         ScrollBars      =   3
         RightMargin     =   10000
         TextRTF         =   $"NILAI.frx":0000
      End
   End
   Begin VB.CommandButton cmdcetak 
      Caption         =   "Cetak"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   12
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdhitung 
      Caption         =   "Hitung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "input data"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2775
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   4695
      Begin VB.TextBox txtuas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   2040
         TabIndex        =   9
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtuts 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   2040
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txttugas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtnama 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Nilai UAS (50%)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Nilai UTS (30%)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Nilai Tugas (20%)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Nama "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Data Nilai Semester Mahasiswa"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8055
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vnama(10), vket(10), vnilai(10) As String
Dim vtugas(10), vuts(10), vuas(10), vrata(10) As Single
Dim i, j As Integer

Private Sub cmdcetak_Click()
Open "D:\daftarnilai.txt" For Output As #1
Print #1, "                    DATA NILAI SEMESTER MAHASISWA"
Print #1, "=============================================================="
Print #1, " NO"; Tab(6); "NAMA"; Tab(21); "RATA2"; Tab(31); "NILAI"; Tab(46); "KET"
Print #1, "=============================================================="

For i = 1 To j
  Print #1, i; Tab(9); vnama(i); Tab(25); vrata(i); Tab(41); vnilai(i); Tab(57); vket(i)
Next i

Print #1, "=============================================================="

Close #1

rtbtampil.FileName = "D:\daftarnilai.txt"

End Sub

Private Sub cmdhitung_Click()
For i = 1 To j
vrata(i) = (20 / 100 * vtugas(i)) + (30 / 100 * vuts(i)) + (50 / 100 * vuas(i))

If vrata(i) >= 80 Then
vket(i) = "LULUS"
 Else
If vrata(i) >= 70 Then
 vket(i) = "LULUS"
 Else
If vrata(i) >= 55 Then
 vket(i) = "LULUS"
 Else
If vrata(i) >= 41 Then
 vket(i) = "GAGAL"
 Else
If vrata(i) < 40 Then
 vket(i) = "GAGAL"
End If
End If
End If
End If
End If

If vrata(i) >= 80 Then
vnilai(i) = "A"
 Else
If vrata(i) >= 70 Then
vnilai(i) = "B"
Else
If vrata(i) >= 55 Then
 vnilai(i) = "C"
 Else
If vrata(i) >= 41 Then
 vnilai(i) = "D"
 Else
If vrata(i) < 40 Then
 vnilai(i) = "E"
End If
End If
End If
End If
End If

Next i

MsgBox "proses hitung selesai", vbOKOnly + vbInformation, "informasi"

End Sub

Private Sub cmdsimpan_Click()
j = j + 1
vnama(j) = txtnama.Text
vtugas(j) = Val(txttugas.Text)
vuts(j) = Val(txtuts.Text)
vuas(j) = Val(txtuas.Text)

txtnama.Text = ""
txttugas.Text = ""
txtuas.Text = ""
txtuts.Text = ""
txtnama.SetFocus
 MsgBox "Data berhasil di simpan.", vbInformation + vbOKOnly, "informasi"

End Sub
