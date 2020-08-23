VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membuat Scrollbar Horizontal di ListBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageByNum _
Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal _
wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

Private Sub Form_Load()
Static x As Long
'Lebar string akan menjadi lebar dari horizontal scroll 'bar tersebut
'Tambahkan suatu string yang panjangnya melebihi lebar 'dari scroll bar yang bersangkutan.
  List1.List(0) = "Selamat datang. Semoga Sukses Menyertai Anda Sekalian!"
  If x < TextWidth(List1.List(0) & " ") Then
     x = TextWidth(List1.List(0) & " ")
     If ScaleMode = vbTwips Then x = x / _
                    Screen.TwipsPerPixelX
        SendMessageByNum List1.hwnd, _
                         LB_SETHORIZONTALEXTENT, x, 0
     End If
End Sub


