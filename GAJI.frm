VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form2 
   Caption         =   "Aplikasi Hitung Gaji"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8760
   LinkTopic       =   "Form2"
   ScaleHeight     =   5160
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "input data"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   840
      TabIndex        =   6
      Top             =   960
      Width           =   5055
      Begin VB.TextBox txtgapok 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   2520
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtnama 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Gaji Pokok"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Karyawan"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "output data"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   720
      TabIndex        =   4
      Top             =   2760
      Width           =   7215
      Begin RichTextLib.RichTextBox rtbtampil 
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2990
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         RightMargin     =   10000
         TextRTF         =   $"GAJI.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdcetak 
      Caption         =   "Cetak"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdhitung 
      Caption         =   "Hitung"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATA GAJI KARYAWAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vnama(10) As String
Dim vgapok(10), vtun(10), vpajak(10) As Single
Dim vgaji(10) As Single
Dim j, i As Integer

Private Sub cmdcetak_Click()
Open "D:\Output.txt" For Output As #1
Print #1, "                       DATA GAJI KARYAWAN"
Print #1, "=============================================================="
Print #1, " NO"; Tab(6); "NAMA"; Tab(21); "GAPOK"; Tab(31); "TUNJANGAN"; Tab(46); "PAJAK"; Tab(57); "GAJI"
Print #1, "=============================================================="

For i = 1 To j
  Print #1, i; Tab(6); vnama(i); Tab(20); vgapok(i); Tab(30); vtun(i); Tab(45); vpajak(i); Tab(55); vgaji(i)
Next i

Print #1, "=============================================================="

Close #1

rtbtampil.FileName = "D:\Output.txt"
End Sub

Private Sub cmdhitung_Click()
For i = 1 To j
vtun(i) = 20 / 100 * vgapok(i)
vpajak(i) = 15 / 100 * (vgapok(i) + vtun(i))
vgaji(i) = vgapok(i) + vtun(i) - vpajak(i)

Next i

MsgBox "proses hitung selesai", vbOKOnly + vbInformation, "informasi"
End Sub

Private Sub cmdsimpan_Click()
j = j + 1
vnama(j) = txtnama.Text
vgapok(j) = Val(txtgapok.Text)

txtnama.Text = ""
txtgapok.Text = ""
txtnama.SetFocus
 MsgBox "Data berhasil di simpan.", vbInformation + vbOKOnly, "informasi"

End Sub

