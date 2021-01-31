VERSION 5.00
Begin VB.Form frmIngreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6540
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   5295
      Begin VB.TextBox txTClave 
         Height          =   495
         Left            =   2280
         TabIndex        =   7
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtUsuario 
         Height          =   495
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CLAVE"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdContinuar 
      Caption         =   "CONTINUAR"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Picture         =   "frmIngreso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "REGRESAR"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      Picture         =   "frmIngreso.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   2
      Top             =   -240
      Width           =   1755
   End
End
Attribute VB_Name = "frmIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
