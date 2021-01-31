VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda"
   ClientHeight    =   3675
   ClientLeft      =   9105
   ClientTop       =   3615
   ClientWidth     =   6615
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid GridSocio 
      Height          =   1215
      Left            =   840
      TabIndex        =   4
      Top             =   1560
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
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
      Height          =   615
      Left            =   4200
      Picture         =   "frmBusqueda.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtBusqueda 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.OptionButton optMovil 
      Caption         =   "MOVIL"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton optCedula 
      Caption         =   "CEDULA"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLSocio As New ADODB.Recordset

Private Sub Form_Load()
    Call ConfigurarGrid
    Call cargargrid
End Sub

Private Sub optCedula_Click()
    Dim sql As String
    sql = "select * from socio where cedula='" & txtBusqueda.Text & "'"
    Set TBLSocio = CONEXION.Execute(sql)
End Sub
Private Sub ConfigurarGrid()
    GridSocio.Clear
    GridSocio.FormatString = "cedula|nombre|apellido"
    GridSocio.ColWidth(0) = 1000
    GridSocio.ColWidth(1) = 1500
    GridSocio.ColWidth(2) = 1000
End Sub
Private Sub cargargrid()
    Dim sql As String
        sql = "select * from socio"
            Set TBLSocio = CONEXION.Execute(sql)
    Dim f As Integer
    f = 1
    GridSocio.Rows = 2
    Do Until TBLSocio.EOF
        GridSocio.TextMatrix(f, 0) = TBLSocio!cedula
        GridSocio.TextMatrix(f, 1) = TBLSocio!nombre
        GridSocio.TextMatrix(f, 2) = TBLSocio!apellido
        
        
    Loop
End Sub
