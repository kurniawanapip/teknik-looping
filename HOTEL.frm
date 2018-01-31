VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form6 
   Caption         =   "Aplikasi Hitung Reservasi Hotel"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form6"
   ScaleHeight     =   6735
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Output Data"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2535
      Left            =   360
      TabIndex        =   13
      Top             =   3840
      Width           =   7455
      Begin RichTextLib.RichTextBox rtbtampil 
         Height          =   2175
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   3836
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         RightMargin     =   10000
         TextRTF         =   $"HOTEL.frx":0000
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
      Left            =   5760
      TabIndex        =   12
      Top             =   2640
      Width           =   1215
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
      Left            =   5760
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
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
      Height          =   555
      Left            =   5760
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Data"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   4815
      Begin MSComCtl2.DTPicker dtpcekout 
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   94502913
         CurrentDate     =   42804
      End
      Begin MSComCtl2.DTPicker dtpcekin 
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   94502913
         CurrentDate     =   42804
      End
      Begin VB.ComboBox cbxkelas 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "HOTEL.frx":0082
         Left            =   2280
         List            =   "HOTEL.frx":008F
         TabIndex        =   7
         Top             =   2040
         Width           =   2295
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
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Kelas Kamar"
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
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Tanggal Cek-Out"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Tanggal Cek-In"
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
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Tamu"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DATA TAMU HOTEL ALEXIS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8175
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vnama(10), vkelas(10) As String
Dim vcekin(10), vcekout(10), vlamainap(10), vbayar(10) As Single
Dim i, j As Integer
Dim vtarif(10) As Single

Private Sub cmdcetak_Click()
Open "D:\tamuhotel.txt" For Output As #1
Print #1, "                          DATA TAMU HOTEL ALEXIS"
Print #1, "========================================================================="
Print #1, " NO"; Tab(6); "NAMA"; Tab(25); "CEK IN"; Tab(40); "CEK OUT"; Tab(55); "KAMAR"; Tab(70); "TARIF"; Tab(80); "HARI"; Tab(90); "BIAYA"
Print #1, "========================================================================="

For i = 1 To j
  Print #1, i; Tab(8); vnama(i); Tab(28); vcekin(i); Tab(43); vcekout(i); Tab(60); vkelas(i); Tab(75); vtarif(i); Tab(88); vlamainap(i); Tab(98); vbayar(i)
Next i

Print #1, "========================================================================="

Close #1

rtbtampil.FileName = "D:\tamuhotel.txt"

End Sub

Private Sub cmdhitung_Click()
For i = 1 To j
If vkelas(i) = "de-lux" Then
vtarif(i) = "400000"
Else
If vkelas(i) = "de-suite" Then
vtarif(i) = "350000"
Else
If vkelas(i) = "vip" Then
vtarif(i) = "500000"
End If
End If
End If

vlamainap(i) = vcekout(i) - vcekin(i)

If vlamainap(i) < 2 Then
vbayar(i) = vtarif(i) * 2
Else
vbayar(i) = vtarif(i) * vlamainap(i)
End If

Next i
MsgBox "proses hitung selesai", vbOKOnly + vbInformation, "informasi"
End Sub

Private Sub cmdsimpan_Click()
j = j + 1
vnama(j) = txtnama.Text
vkelas(j) = cbxkelas.Text
vcekin(j) = dtpcekin.Value
vcekout(j) = dtpcekout.Value

txtnama.Text = ""
cbxkelas.Text = ""
txtnama.SetFocus
 MsgBox "Data berhasil di simpan.", vbInformation + vbOKOnly, "informasi"

End Sub

