VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7935
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAcerca 
      Caption         =   "Acerca"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6360
      Picture         =   "frmMenu.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1450
   End
   Begin VB.CommandButton cmdReemprimirAportacion 
      Caption         =   "Reimprimir Aportacion"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6360
      Picture         =   "frmMenu.frx":1BEA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   1450
   End
   Begin VB.CommandButton cmbReenvioMULTAS 
      Caption         =   "Reimprimir Multas"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6360
      Picture         =   "frmMenu.frx":258D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gestión de Multas"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      Picture         =   "frmMenu.frx":2F30
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1450
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Gestión de Aportaciones"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      Picture         =   "frmMenu.frx":618A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1450
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Gestión de Socios"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6360
      Picture         =   "frmMenu.frx":9D29
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1450
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cobro de Multas"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      Picture         =   "frmMenu.frx":CF15
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1450
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cobro de aportaciones"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      Picture         =   "frmMenu.frx":10219
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1450
   End
   Begin VB.Image Image1 
      Height          =   4695
      Left            =   1440
      Picture         =   "frmMenu.frx":13402
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbReenvioMULTAS_Click()
    FrmReimpresionM.Show
    frmMenu.Hide
End Sub

Private Sub cmdAcerca_Click()
    frmAcercaDE.Show
    frmMenu.Hide
End Sub

Private Sub cmdReemprimirAportacion_Click()
    frmReemprisionAportacion.Show
    frmMenu.Hide
End Sub

Private Sub Command1_Click()
    frmGestionMultas.Show
    frmMenu.Hide

End Sub
Private Sub Command2_Click()
    frmAportaciones.Show
    frmMenu.Hide
End Sub
Private Sub Command3_Click()
    frmSocios.Show
    frmMenu.Hide
End Sub
Private Sub Command4_Click()
    FrmCobranzasMultas.Show
    frmMenu.Hide
End Sub
Private Sub Command5_Click()
    FrmCobranzasAportaciones.Show
    frmMenu.Hide
End Sub
