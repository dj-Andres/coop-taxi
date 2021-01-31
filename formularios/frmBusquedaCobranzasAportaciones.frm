VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBusquedaCobranzasAportaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Cobranzas Aportaciones"
   ClientHeight    =   3375
   ClientLeft      =   11475
   ClientTop       =   3150
   ClientWidth     =   6630
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmBusquedaCobranzasAportaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBusqueda 
      Height          =   375
      Left            =   600
      MaxLength       =   15
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin MSFlexGridLib.MSFlexGrid GridBusquedaCobranzas 
      Height          =   1575
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      FixedCols       =   0
      FocusRect       =   0
      ScrollBars      =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmBusquedaCobranzasAportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cisoci As String
Dim IDSOCIO As Integer
Dim TBLCobrar As New ADODB.Recordset
Private Sub Form_Load()
    ModuloBaseDatos.conectardb
    'txtBusqueda.SetFocus
    Call configurargrid
    Call cargargrid
End Sub
Private Sub configurargrid()
    GridBusquedaCobranzas.Clear
    GridBusquedaCobranzas.FormatString = "cedula|nombre|apellido"
    GridBusquedaCobranzas.ColWidth(0) = 1000
    GridBusquedaCobranzas.ColWidth(1) = 1500
    GridBusquedaCobranzas.ColWidth(2) = 1800
    'GridBusquedaCobranzas.ColWidth(3) = 1800
    'GridBusquedaCobranzas.ColWidth(4) = 1500
    'GridBusquedaCobranzas.ColWidth(5) = 1500

End Sub
Private Sub cargargrid()
    Dim sql As String
            sql = "select  * from socio"
            Set TBLCobrar = CONEXION.Execute(sql)
    Dim f As Integer
    f = 1
    GridBusquedaCobranzas.Rows = 2
    Do Until TBLCobrar.EOF
        GridBusquedaCobranzas.TextMatrix(f, 0) = TBLCobrar!cedula
        GridBusquedaCobranzas.TextMatrix(f, 1) = TBLCobrar!nombre
        GridBusquedaCobranzas.TextMatrix(f, 2) = TBLCobrar!apellido
        'GridBusquedaCobranzas.TextMatrix(f, 3) = TBLCobrar!nombre_multa
        'GridBusquedaCobranzas.TextMatrix(f, 4) = TBLCobrar!nombre_multa
        'GridBusquedaCobranzas.TextMatrix(f, 5) = TBLCobrar!fecha
       
        TBLCobrar.MoveNext
        f = f + 1
        GridBusquedaCobranzas.Rows = GridBusquedaCobranzas.Rows + 1
        
    Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmCobranzasAportaciones.Enabled = True
Unload Me
End Sub

Private Sub GridBusquedaCobranzas_DblClick()
    Dim z As Integer
    z = GridBusquedaCobranzas.Row
    If z > 0 Then
        cisocio = GridBusquedaCobranzas.TextMatrix(GridBusquedaCobranzas.Row, 0)
        idsociomulta = ModuloFunciones.buscarID("socio", "cedula", cisocio)
        IDSOCIO1 = idsociomulta
        Set TBLCobrar = Nothing
        TBLCobrar.Open "select * from socio where cedula= '" & cisocio & "'", CONEXION, adOpenDynamic, adLockOptimistic
    

        TBLCobrar.MoveFirst
        
        FrmCobranzasAportaciones.txtcedula.Text = TBLCobrar.Fields("cedula").Value
        FrmCobranzasAportaciones.txtnombre.Text = TBLCobrar.Fields("nombre").Value
        FrmCobranzasAportaciones.txtapellido.Text = TBLCobrar.Fields("apellido").Value
        'FrmCobranzasMultas.txtValor.Text = TBLCobrar.Fields("valor").Value
        'FrmCobranzasMultas.txtsaldo.Text = TBLCobrar.Fields("saldo").Value
    Else
    End If
        Unload Me
        FrmCobranzasAportaciones.Enabled = True
        
    

End Sub
Private Sub txtBusqueda_Change()
 
    Dim sql As String
    Set TBLSocio = Nothing
    sql = "select * from socio where cedula like '%" & txtBusqueda & "%' or nombre like '%" & Trim(UCase(txtBusqueda.Text)) & "%' or apellido like '%" & Trim(UCase(txtBusqueda.Text)) & "%'"
    Set TBLCobrar = CONEXION.Execute(sql)
    
    Dim f As Integer
    f = 1
    GridBusquedaCobranzas.Rows = 2
        Do Until TBLCobrar.EOF
            GridBusquedaCobranzas.TextMatrix(f, 0) = TBLCobrar!cedula
            GridBusquedaCobranzas.TextMatrix(f, 1) = TBLCobrar!nombre
            GridBusquedaCobranzas.TextMatrix(f, 2) = TBLCobrar!apellido
             'GridBusquedaCobranzas.TextMatrix(f, 3) = TBLCobrar!apellido
             'GridBusquedaCobranzas.TextMatrix(f, 4) = TBLCobrar!nombre_multa
             'GridBusquedaCobranzas.TextMatrix(f, 5) = TBLCobrar!fecha
            TBLCobrar.MoveNext
            f = f + 1
            GridBusquedaCobranzas.Rows = GridBusquedaCobranzas.Rows + 1
            
        Loop

End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
KeyAscii = ModuloFunciones.numeros_letras(KeyAscii)
End Sub
