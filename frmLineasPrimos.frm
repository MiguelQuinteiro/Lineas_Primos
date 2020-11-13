VERSION 5.00
Begin VB.Form frmLineasPrimos 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   Caption         =   "Lineas Primos"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   15255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Controles "
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   8
         Text            =   "250"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtOrigenY 
         Height          =   495
         Left            =   1560
         TabIndex        =   6
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtOrigenX 
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdIzquierda 
         Caption         =   "Izquierda"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdDerecha 
         Caption         =   "Derecha"
         Height          =   495
         Left            =   1560
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdAbajo 
         Caption         =   "Abajo"
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdArriba 
         Caption         =   "Arriba"
         Height          =   495
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmLineasPrimos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim miOrigenX As Long
Dim miOrigenY As Long
Dim miIncremento As Long
Dim miAncho As Long
Dim miTurno As Long

Private Sub Form_Load()
  miOrigenX = frmLineasPrimos.Width / 2
  miOrigenY = frmLineasPrimos.Height / 2
  txtOrigenX.Text = miOrigenX
  txtOrigenY.Text = miOrigenY
  miAncho = 1
  miIncremento = 100
End Sub

Private Sub cmdArriba_Click()
  Line (miOrigenX, miOrigenY)-(miOrigenX, miOrigenY - miAncho)
  miOrigenY = miOrigenY - miAncho
  'miAncho = miAncho + miIncremento
End Sub

Private Sub cmdAbajo_Click()
  Line (miOrigenX, miOrigenY)-(miOrigenX, miOrigenY + miAncho)
  miOrigenY = miOrigenY + miAncho
  'miAncho = miAncho + miIncremento
End Sub

Private Sub cmdIzquierda_Click()
  Line (miOrigenX, miOrigenY)-(miOrigenX - miAncho, miOrigenY)
  miOrigenX = miOrigenX - miAncho
  'miAncho = miAncho + miIncremento
End Sub

Private Sub cmdDerecha_Click()
  Line (miOrigenX, miOrigenY)-(miOrigenX + miAncho, miOrigenY)
  miOrigenX = miOrigenX + miAncho
  'miAncho = miAncho + miIncremento
End Sub

Private Sub cmdMostrar_Click()
' Limpia la pantalla
  Cls

  ' Inicializa en el centro de la pantalla
  miOrigenX = frmLineasPrimos.Width / 2
  miOrigenY = frmLineasPrimos.Height / 2
  txtOrigenX.Text = miOrigenX
  txtOrigenY.Text = miOrigenY

  ' Inicializa el ancho
  miAncho = 1
  miTurno = 0

  Dim i As Long
  For i = 1 To (Val(txtN.Text))
    miTurno = miTurno + 1

    If Primo(i) Then
      miAncho = miAncho + (i * 1.1)
      frmLineasPrimos.ForeColor = vbRed
    Else
      'frmLineasPrimos.ForeColor = vbBlue
    End If

    ' Mueve arriba
    If miTurno Mod 4 = 0 Then
      Call cmdArriba_Click
    End If
    ' Mueve Izquierda
    If miTurno Mod 4 = 1 Then
      Call cmdIzquierda_Click
    End If
    ' Mueve Abajo
    If miTurno Mod 4 = 2 Then
      Call cmdAbajo_Click
    End If
    ' Mueve Derecha
    If miTurno Mod 4 = 3 Then
      Call cmdDerecha_Click
    End If

  Next i


End Sub



' FUNCION PARA CALCULAR SI EL NUMERO ES PRIMO
Public Function Primo(ByVal pN As Long) As Boolean
  Dim i As Long
  Primo = True
  If pN = 1 Then
    Primo = False
  Else
    For i = 2 To Sqr(pN)
      If (pN / i) = Int(pN / i) Then
        Primo = False
      End If
    Next i
  End If
End Function

