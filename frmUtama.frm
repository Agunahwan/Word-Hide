VERSION 5.00
Begin VB.Form frmUtama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Word Hide"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   Icon            =   "frmUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtJawab 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   0
      Top             =   4440
      Width           =   2055
   End
   Begin VB.PictureBox pctKalimat 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   11235
      TabIndex        =   1
      Top             =   2520
      Width           =   11295
      Begin VB.Frame fraKalimat 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   7335
         Begin VB.Line lnKalimat 
            Index           =   0
            X1              =   240
            X2              =   600
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label lblKalimat 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   375
         End
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   3360
      X2              =   11400
      Y1              =   710
      Y2              =   710
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   3360
      X2              =   11400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Word Hide v1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2010"
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
      Left            =   3360
      TabIndex        =   9
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed By Agunahwan Absin"
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
      Left            =   3360
      TabIndex        =   8
      Top             =   840
      Width           =   4935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   11400
      Y1              =   1545
      Y2              =   1545
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   11400
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Image imgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   120
      Picture         =   "frmUtama.frx":6852
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblKesempatan 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9600
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "K E S E M P A T A N :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label lblSkor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "N I L A I :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image imgBackground 
      Height          =   5415
      Left            =   0
      Picture         =   "frmUtama.frx":CBB9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#########################################################################
'###########    Nama       : Word Hide          ##########################
'###########    Version    : 1.0                ##########################
'###########    Programmer : Agunahwan Absin    ##########################
'###########    Copyright (c) 2010              ##########################
'#########################################################################

Option Explicit

Dim Jumlah As Integer
Dim Kalimat As String
Dim JumlahTebakan As Integer

'Variabel untuk memanggil Database
Dim Nomor(10000) As Long
Dim Word(10000) As String

Sub Inisial()
Dim i, Cari As Integer
Dim Baris As String
    'Membuat background
    imgBackground.Width = Me.Width
    imgBackground.Height = Me.Height
    
    'Memasukkan database variabel
    Open App.Path & "\kata.db" For Input As #1
    Jumlah = LOF(1)
    i = 1
    Do While Not EOF(1)
        Line Input #1, Baris
        Baris = Trim(Baris)
        Cari = InStr(1, Baris, ":")
        Nomor(i) = CLng(Left(Baris, Cari - 1))
        Word(i) = CStr(Right(Baris, Len(Baris) - Cari))
        i = i + 1
    Loop
    Jumlah = i - 2
    Close #1
End Sub

Sub BuatKalimat()
Dim Min, Max, Kode As Integer
Dim i As Integer
    Max = Jumlah
    Min = 1
    Randomize Timer
    Kode = Int((Max - Min + 1) * Rnd) + Min
    
    'Inisial Kesempatan
    lblKesempatan.Caption = "10"
    
    'Inisial Jumlah Tebakan
    JumlahTebakan = 0
    
    'Kalimat terpilih
    Kalimat = UCase(Word(Kode))
    
    'Inisialisasi Frame
    fraKalimat.Width = ((lblKalimat(0).Width + 250) * (Len(Kalimat) + 1)) + lblKalimat(0).Width
    fraKalimat.Left = (pctKalimat.Width - fraKalimat.Width) / 2
    
    'Memasukkan setiap kata ke tag label
    lblKalimat(0).Tag = Left(Kalimat, 1)
    lblKalimat(0).Left = (lblKalimat(0).Width + 250)
    lnKalimat(0).X1 = lblKalimat(0).Left
    lnKalimat(0).X2 = lnKalimat(0).X1 + 360
    
    For i = 1 To Len(Kalimat) - 1
        Load lblKalimat(i)
        lblKalimat(i).Visible = True
        lblKalimat(i).Left = (lblKalimat(0).Width + 250) * (i + 1)
        lblKalimat(i).Tag = Mid(Kalimat, i + 1, 1)
    
        Load lnKalimat(i)
        lnKalimat(i).Visible = True
        lnKalimat(i).X1 = lblKalimat(i).Left
        lnKalimat(i).X2 = lnKalimat(i).X1 + 360
    Next
End Sub

Sub BukaKalimat()
Dim i As Integer
    For i = 0 To Len(Kalimat) - 1
        lblKalimat(i).Caption = lblKalimat(i).Tag
    Next
End Sub

Sub BuangObjek()
Dim i As Integer
    For i = 1 To Len(Kalimat) - 1
        Unload lblKalimat(i)
        Unload lnKalimat(i)
    Next
    lblKalimat(0).Caption = ""
End Sub

Sub Periksa(Huruf As String)
Dim i As Integer
Dim Ada As Boolean
    Ada = False
    For i = 0 To Len(Kalimat) - 1
        If lblKalimat(i).Tag = Huruf And lblKalimat(i).Caption <> Huruf Then
            lblKalimat(i).Caption = lblKalimat(i).Tag
            Ada = True
            JumlahTebakan = JumlahTebakan + 1
            
            'Menambahkan nilai skor
            lblSkor.Caption = CStr(CInt(lblSkor.Caption) + 10)
        End If
    Next
    
    'Periksa jika huruf tidak ada maka jumlah kesempatan dikurangi satu
    'dan nilai dikurangi lima
    If Ada = False Then
        lblSkor.Caption = CStr(CInt(lblSkor.Caption) - 5)
        If lblKesempatan <> "1" Then
            lblKesempatan.Caption = CStr(CInt(lblKesempatan.Caption) - 1)
        Else
            BukaKalimat
            lblKesempatan.Caption = "0"
            MsgBox "Kesempatan Anda habis. Kata tersembunyi adalah " & Kalimat & ".", vbExclamation + vbOKOnly, "Selesai"
            BuangObjek
            BuatKalimat
        End If
    Else
        If JumlahTebakan = Len(Kalimat) Then
            txtJawab.Text = ""
            MsgBox "Selamat, semua huruf telah terbuka. Kata tersembunyi adalah " & Kalimat & ".", vbOKOnly + vbInformation, "Selesai"
            BuangObjek
            BuatKalimat
        End If
    End If
End Sub

Private Sub Form_Load()
    Inisial
    BuatKalimat
End Sub

Private Sub txtJawab_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtJawab.Text) <> "" Then
        If KeyCode = 13 Then
            Periksa UCase(txtJawab.Text)
            txtJawab.Text = ""
        End If
    Else
        txtJawab.Text = ""
    End If
End Sub
