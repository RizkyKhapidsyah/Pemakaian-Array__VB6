VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCari 
      Caption         =   "Cari"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdTampil 
      Caption         =   "Tampil"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTampil_Click()
Dim i As Integer
  'Tampilkan isi array tabData
  For i = 7 To 0 Step -1
      MsgBox tabData(i).Isi 'Tampilkan ke layar satu per satu
  Next i
End Sub

Private Sub cmdCari_Click()
Dim kriteria As String
Dim urut As Integer
Dim ketemu As Boolean
  'Tampung data yang akan dicari
  kriteria = InputBox("Masukkan data yang akan dicari (1 karakter)")
  If kriteria = "" Then Exit Sub  'Jika kosong, langsung keluar
  ketemu = False 'Inisialisasi ketemu (masih belum ketemu)
  urut = 0 'Inisialisasi utk mengetahui posisi sebenarnya
  For i = 7 To 0 Step -1
      If tabData(i).Isi = kriteria Then 'Jika ditemukan
         MsgBox "Data '" & kriteria & "' ditemukan setelah dibalik" & Chr(13) & _
                "berada pada urutan ke-" & urut + 1 & "", vbInformation
         ketemu = True 'berarti sudah pernah ketemu
      End If
      urut = urut + 1 'untuk mencari posisi dari awal
  Next i
  If ketemu = False Then
     'Jika tidak ditemukan, tampilkan pesan
     MsgBox "Data " & kriteria & " tidak ditemukan!", vbCritical
  Else 'jika sudah pernah ketemu, langsung keluar
     Exit Sub
  End If
End Sub

Private Sub Form_Load()
ReDim tabData(8) 'Isi array sebanyak delapan elemen
  'Isi array tabData
  tabData(0).Isi = "h"
  tabData(1).Isi = "a"
  tabData(2).Isi = "c"
  tabData(3).Isi = "k"
  tabData(4).Isi = "1"
  tabData(5).Isi = "4"
  tabData(6).Isi = "1"
  tabData(7).Isi = "2"
End Sub

