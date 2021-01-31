VERSION 5.00
Begin VB.Form frmConsultasMultas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONSULTAS MULTAS"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7440
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optMultas 
      Caption         =   "MULTAS"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Width           =   2655
   End
   Begin VB.OptionButton optFechaCobro 
      Caption         =   "FECHA DE COBRO"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.OptionButton optEstadoPago 
      Caption         =   "ESTADO DE PAGO"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmConsultasMultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub optEstadoPago_Click()
    frmEstadoPagos.Show
End Sub
Private Sub optFechaCobro_Click()
    frmFechaCobro.Show
End Sub
Private Sub optMultas_Click()
    frmMultasSocios.Show
End Sub
