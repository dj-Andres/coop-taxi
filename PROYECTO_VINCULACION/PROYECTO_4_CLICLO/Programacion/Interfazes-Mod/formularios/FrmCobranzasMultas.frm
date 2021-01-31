VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCobranzasMultas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cobranzas  Multas"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8820
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "FrmCobranzasMultas.frx":0000
   ScaleHeight     =   7170
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuscar 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Picture         =   "FrmCobranzasMultas.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1815
      Left            =   600
      TabIndex        =   9
      Top             =   3720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   10
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtValor 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtSocio 
      Height          =   405
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   110952449
      CurrentDate     =   43348
   End
   Begin VB.Label LblImprimir 
      Alignment       =   2  'Center
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   6600
      Picture         =   "FrmCobranzasMultas.frx":0C0C
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SALDO"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   8
      Top             =   3240
      Width           =   690
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   6
      Top             =   2640
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SOCIO"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label lblCobranzasMultas 
      Alignment       =   2  'Center
      Caption         =   "Cobranzas de Multas"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "FrmCobranzasMultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    frmBusqueda.Show
End Sub
