VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCobranzasAporatciones 
   Caption         =   "Cobranzas de Aportaciones"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   4320
      Picture         =   "FrmCobranzasAporatciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2055
      Left            =   720
      TabIndex        =   7
      Top             =   3600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   10
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.TextBox txtValor 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txtSocio 
      Height          =   405
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   110952449
      CurrentDate     =   43348
   End
   Begin VB.Label lbltotalPagar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL PAGAR"
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
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   1380
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
      Left            =   480
      TabIndex        =   4
      Top             =   2400
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
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label lblCobranzasAportaciones 
      Alignment       =   2  'Center
      Caption         =   "Cobranzas de Aportaciones"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
End
Attribute VB_Name = "FrmCobranzasAporatciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuscar_Click()

End Sub

Private Sub Command1_Click()
    frmBusqueda.Show
End Sub
