VERSION 5.00
Begin VB.Form frmAcercaDE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9765
   Icon            =   "frmAcercaDE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   7440
         TabIndex        =   1
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Versión instalada: 18.1.1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2880
         TabIndex        =   9
         Top             =   1320
         Width           =   3885
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema de Pagos de Multas Y Aportaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SystFinCont"
         BeginProperty Font 
            Name            =   "Imprint MT Shadow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   1905
      End
      Begin VB.Image Image3 
         Height          =   1215
         Left            =   480
         Picture         =   "frmAcercaDE.frx":08CA
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   1695
         Left            =   0
         Top             =   120
         Width           =   8730
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   240
         X2              =   8280
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label4 
         Caption         =   "Carrera de Tecnología en Análisis de Sistemas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3120
         Width           =   4935
      End
      Begin VB.Label Label3 
         Caption         =   "Sistema informático desarrollado como proyecto de Vinculación, entre el Instituto y la empresa Cooperativa de Taxis 15 de Octubre."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   6015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "https://itsjol.edu.ec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   3480
         Width           =   1650
      End
      Begin VB.Line Line2 
         X1              =   -120
         X2              =   8520
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label7 
         Caption         =   "Copyright © 2018 Instituto Tecnológico Superior José Ochoa León. Todos los derechos reservados."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   6015
      End
      Begin VB.Image Image2 
         Height          =   1800
         Left            =   6360
         Picture         =   "frmAcercaDE.frx":4E097
         Top             =   1920
         Width           =   1800
      End
      Begin VB.Label Label5 
         Caption         =   $"frmAcercaDE.frx":504E6
         Height          =   585
         Left            =   240
         TabIndex        =   2
         Top             =   4080
         Width           =   7215
      End
   End
End
Attribute VB_Name = "frmAcercaDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    frmMenu.Show
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMenu.Show
    Unload Me
End Sub
