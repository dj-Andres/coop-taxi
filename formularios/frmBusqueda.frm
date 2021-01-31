VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda"
   ClientHeight    =   3675
   ClientLeft      =   9570
   ClientTop       =   1770
   ClientWidth     =   6615
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmBusqueda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid GridBusqueda 
      Height          =   1575
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2778
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      ScrollBars      =   0
      SelectionMode   =   1
   End
   Begin VB.TextBox txtBusqueda 
      Height          =   375
      Left            =   840
      MaxLength       =   15
      TabIndex        =   0
      Top             =   840
      Width           =   4815
   End
End
Attribute VB_Name = "frmBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cisocio As String
Dim IDSOCIO As Integer
Dim TBLSocio As New ADODB.Recordset
Private Sub configurargrid()
                                                                                                                                                     
    GridBusqueda.Clear
    GridBusqueda.FormatString = "movil|cedula|nombre|apellido"
    GridBusqueda.ColWidth(0) = 1000
    GridBusqueda.ColWidth(1) = 1500
    GridBusqueda.ColWidth(2) = 1500
    GridBusqueda.ColWidth(3) = 1500


End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ModuloBaseDatos.conectardb
    Call configurargrid
    Call cargargrid
End Sub
Private Sub cargargrid()
    Dim sql As String
    sql = "select * from socio"
        Set TBLSocio = CONEXION.Execute(sql)
Dim f As Integer
f = 1
GridBusqueda.Rows = 2
    Do Until TBLSocio.EOF
        GridBusqueda.TextMatrix(f, 0) = TBLSocio!movil
        GridBusqueda.TextMatrix(f, 1) = TBLSocio!cedula
        GridBusqueda.TextMatrix(f, 2) = TBLSocio!nombre
        GridBusqueda.TextMatrix(f, 3) = TBLSocio!apellido
        TBLSocio.MoveNext
        f = f + 1
        GridBusqueda.Rows = GridBusqueda.Rows + 1
        
    Loop

End Sub

Private Sub GridBusqueda_Click()
    Dim z As Integer
    z = GridBusqueda.Row
    If z > 0 Then
        cisocio = GridBusqueda.TextMatrix(GridBusqueda.Row, 1)
        IDSOCIO = ModuloFunciones.buscarID("socio", "cedula", cisocio)
        
        Set TBLSocio = Nothing
        TBLSocio.Open "select * from socio where cedula='" & cisocio & "'", CONEXION, adOpenDynamic, adLockOptimistic
        If Not (TBLSocio.EOF) Then
        TBLSocio.MoveFirst
        
        frmGestionMultas.txtSocio.Text = TBLSocio.Fields("cedula").Value
        frmGestionMultas.txtnombre.Text = TBLSocio.Fields("nombre").Value
        frmGestionMultas.txtapellido.Text = TBLSocio.Fields("apellido").Value
    Else
    End If
        Unload Me
        frmGestionMultas.Enabled = True
        
    End If
    
End Sub
Private Sub txtBusqueda_Change()
    Dim sql As String
    Set TBLSocio = Nothing
    sql = "select * from socio where cedula like '%" & txtBusqueda.Text & "%'  or  nombre like '%" & Trim(UCase(txtBusqueda.Text)) & "%'  "
    Set TBLSocio = CONEXION.Execute(sql)
    
    Dim f As Integer
    f = 1
    GridBusqueda.Rows = 2
        Do Until TBLSocio.EOF
            GridBusqueda.TextMatrix(f, 0) = TBLSocio!movil
            GridBusqueda.TextMatrix(f, 1) = TBLSocio!cedula
            GridBusqueda.TextMatrix(f, 2) = TBLSocio!nombre
            GridBusqueda.TextMatrix(f, 3) = TBLSocio!apellido
            TBLSocio.MoveNext
            f = f + 1
            GridBusqueda.Rows = GridBusqueda.Rows + 1
            
        Loop

    
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
KeyAscii = ModuloFunciones.numeros_letras(KeyAscii)
End Sub
