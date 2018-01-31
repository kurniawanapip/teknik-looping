VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form3 
   Caption         =   "Aplikasi Hitung Tagihan PLN"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8685
   LinkTopic       =   "Form3"
   ScaleHeight     =   6630
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FFFF&
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
      Height          =   3015
      Left            =   600
      TabIndex        =   11
      Top             =   3360
      Width           =   7455
      Begin RichTextLib.RichTextBox rtbtampil 
         Height          =   2655
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4683
         _Version        =   393217
         ScrollBars      =   3
         RightMargin     =   10000
         TextRTF         =   $"PLN.frx":0000
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
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdhitung 
      Caption         =   "Hitung"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Data"
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
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   4935
      Begin VB.ComboBox cbxdaya 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "PLN.frx":0080
         Left            =   2520
         List            =   "PLN.frx":008D
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtkwh 
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
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtnama 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Pemakaian KWH"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Besar Daya"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Pelanggan"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Tagihan PLN (Listrik Untuk Hidup Lebih Baik)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vnama(10) As String
Dim vkwh(10), vdaya(10), vabod(10), vpajak(10), vpakai(10) As Single
Dim vtarif(20) As Single
Dim vtagihan(10) As Single
Dim j, i As Integer

Private Sub cmdcetak_Click()
Open "D:\2116007.txt" For Output As #1
Print #1, "                          DATA TAGIHAN PELANGGAN PLN"
Print #1, "==========================================================================================="
Print #1, " NO"; Tab(6); "NAMA"; Tab(20); "DAYA"; Tab(30); "KWH"; Tab(45); "ABODEMEN"; Tab(60); "BY.KWH"; Tab(70); "PPJ"; Tab(80); "TAGIHAN"
Print #1, "==========================================================================================="

For i = 1 To j
  Print #1, i; Tab(6); vnama(i); Tab(19); vdaya(i); Tab(29); vkwh(i); Tab(45); vabod(i); Tab(59); vpakai(i); Tab(69); vpajak(i); Tab(80); vtagihan(i)
Next i

Print #1, "==========================================================================================="

Close #1

rtbtampil.FileName = "D:\2116007.txt"

End Sub

Private Sub cmdhitung_Click()
For i = 1 To j

 If vdaya(i) = 450 Then
 vtarif(i) = "250"
 Else
 If vdaya(i) = 900 Then
 vtarif(i) = "350"
 Else
 If vdaya(i) = 1300 Then
 vtarif(i) = "550"
 End If
 End If
 End If
 
 If vdaya(i) = 1300 Then
  vabod(i) = "35000"
 Else
  vabod(i) = "20000"
 End If
 
 vpakai(i) = vtarif(i) * vkwh(i)
 vpajak(i) = 5 / 100 * vpakai(i)
 vtagihan(i) = vabod(i) + vpakai(i) + vpajak(i)
Next i

MsgBox "proses hitung selesai", vbOKOnly + vbInformation, "informasi"
End Sub

Private Sub cmdsimpan_Click()
j = j + 1
vnama(j) = txtnama.Text
vdaya(j) = Val(cbxdaya.Text)
vkwh(j) = Val(txtkwh.Text)

txtnama.Text = ""
txtkwh.Text = ""
cbxdaya.Text = ""
txtnama.SetFocus
MsgBox "data berhasil di input", vbOKOnly + vbInformation, "informasi"
End Sub
