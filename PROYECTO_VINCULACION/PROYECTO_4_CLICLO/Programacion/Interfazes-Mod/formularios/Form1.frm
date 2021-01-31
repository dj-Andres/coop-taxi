VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGestionMultas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multas"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9030
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmdbuscar 
      Height          =   615
      Left            =   4080
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Txtbuscar 
      Height          =   495
      Left            =   1680
      TabIndex        =   19
      Top             =   3720
      Width           =   2295
   End
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
      Left            =   4680
      Picture         =   "Form1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   840
      Width           =   615
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
      Left            =   7200
      Picture         =   "Form1.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   6855
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
         Picture         =   "Form1.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1335
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
         Picture         =   "Form1.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1455
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
         Picture         =   "Form1.frx":2BF2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
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
         Picture         =   "Form1.frx":34BC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridMultas 
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   5
      Cols            =   5
      FixedCols       =   0
   End
   Begin VB.ComboBox cmbMultas 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   1080
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      Format          =   109838337
      CurrentDate     =   43348
   End
   Begin VB.TextBox txtValor 
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtObservacion 
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   4695
   End
   Begin VB.TextBox txtSocio 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Lblbuscar 
      Caption         =   "BUSCAR"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   3720
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   240
      Picture         =   "Form1.frx":3D86
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label Label7 
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
      Left            =   120
      TabIndex        =   16
      Top             =   7320
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4680
      Picture         =   "Form1.frx":4650
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label5 
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
      Left            =   600
      TabIndex        =   6
      Top             =   2880
      Width           =   675
   End
   Begin VB.Label Label4 
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
      Left            =   600
      TabIndex        =   5
      Top             =   2400
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVACION"
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
      TabIndex        =   3
      Top             =   1800
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MULTA"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   690
   End
   Begin VB.Label Label1 
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
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   645
   End
End
Attribute VB_Name = "frmGestionMultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLGuardarMulta As New ADODB.Recordset
Private Sub cmbMultas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtObservacion.SetFocus
End Sub
Private Sub CmdAgregar1_Click()
    FrmAgregarMultas.Show
End Sub
Private Sub CmdGuardar_Click()
    Dim respuestaID As Integer
    TBLGuardarMulta = Nothing
    set tblguardarmulta.Open "select * from socio_multas where idsocio = '" & txtsocio.Text & "'", CONEXION, adOpenDynamic, adLockOptimistic
    
    
    
End Sub

Private Sub DTPFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtValor.SetFocus
End Sub
Dim TBLGuardarMulta As New ADODB.Recordset
Private Sub Form_Load()
    ModuloBaseDatos.conectardb
    Call ConfigurarGrid
    Call cargargrid
End Sub
Private Sub Image1_Click()
    frmBusqueda.Show
End Sub
Private Sub ConfigurarGrid()
    GridMultas.Clear
    GridMultas.FormatString = "socio|multa|observacion|fecha|valor"
    GridMultas.ColWidth(0) = 1000
    GridMultas.ColWidth(1) = 1500
    GridMultas.ColWidth(2) = 1500
    GridMultas.ColWidth(3) = 1500
    GridMultas.ColWidth(4) = 1500
    
End Sub
Private Sub cargargrid()
    Dim sql As String
            sql = "select * from v_socio_multas"
                Set TBLMulta = CONEXION.Execute(sql)
        Dim f As Integer
        f = 1
        GridMultas.Rows = 2
        Do Until TBLMulta.EOF
            GridMultas.TextMatrix(f, 0) = TBLMulta!nombre
            GridMultas.TextMatrix(f, 1) = TBLMulta!nombre_multa
            GridMultas.TextMatrix(f, 2) = TBLMulta!observacion
            GridMultas.TextMatrix(f, 3) = TBLMulta!fecha
            GridMultas.TextMatrix(f, 4) = TBLMulta!valor
            TBLMulta.MoveNext
            f = f + 1
            GridMultas.Rows = GridMultas.Rows + 1
            
        Loop

End Sub
Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.DTPFecha.SetFocus
End Sub

Private Sub txtSocio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbMultas.SetFocus
        KeyAscii = ModuloFunciones.letras(KeyAscii)
End Sub
