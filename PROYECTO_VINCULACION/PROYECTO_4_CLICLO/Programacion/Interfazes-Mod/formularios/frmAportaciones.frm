VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAportaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aportaciones"
   ClientHeight    =   9975
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9525
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAgregar1 
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
      Left            =   4440
      Picture         =   "frmAportaciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1200
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   240
      TabIndex        =   18
      Top             =   6960
      Width           =   6855
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "GUARDAR"
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
         Left            =   1920
         Picture         =   "frmAportaciones.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "AGREGAR"
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
         Left            =   240
         Picture         =   "frmAportaciones.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "MODIFICAR"
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
         Picture         =   "frmAportaciones.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "CANCELAR"
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
         Left            =   5160
         Picture         =   "frmAportaciones.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "CERRAR"
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
      Left            =   7440
      Picture         =   "frmAportaciones.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7200
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1215
      Left            =   600
      TabIndex        =   16
      Top             =   5400
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   8
      Cols            =   8
      FixedCols       =   0
   End
   Begin VB.TextBox txtValor 
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txtHora 
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   3960
      Width           =   1695
   End
   Begin VB.ComboBox cmbEstado 
      Height          =   315
      Left            =   2400
      TabIndex        =   11
      Top             =   3480
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   110952449
      CurrentDate     =   43348
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   2040
      Width           =   5775
   End
   Begin VB.ComboBox cmbAportaciones 
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtSocio 
      Height          =   405
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtCodigo 
      Height          =   405
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4080
      Picture         =   "frmAportaciones.frx":34BC
      Stretch         =   -1  'True
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENVIAR A IMPRIMIR"
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
      Left            =   360
      TabIndex        =   23
      Top             =   8640
      Width           =   1890
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   600
      Picture         =   "frmAportaciones.frx":3D86
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   1215
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
      Left            =   480
      TabIndex        =   15
      Top             =   4680
      Width           =   675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HORA"
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
      TabIndex        =   13
      Top             =   4080
      Width           =   570
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Left            =   480
      TabIndex        =   10
      Top             =   3480
      Width           =   1770
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
      TabIndex        =   8
      Top             =   2760
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION"
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
      TabIndex        =   7
      Top             =   2160
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO APORTACION"
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
      Top             =   1440
      Width           =   1830
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
      TabIndex        =   3
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
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
      TabIndex        =   1
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "frmAportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAgregar1_Click()
    FrmAgregarAportaciones.Show
End Sub

Private Sub Image1_Click()
    frmBusqueda.Show
End Sub
