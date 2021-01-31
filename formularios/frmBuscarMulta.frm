VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBuscarMulta 
   Caption         =   "Buscar Multas"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtBusqueda 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin MSFlexGridLib.MSFlexGrid GridBusquedaMulta 
      Height          =   1575
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   4
      Cols            =   6
      FixedCols       =   0
      FocusRect       =   0
      ScrollBars      =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmBuscarMulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cimulta As Integer
Dim idmulta As Integer
Dim TBLMulta As New ADODB.Recordset
Private Sub Form_Load()
    ModuloBaseDatos.conectardb
    Call configurargrid
    Call cargargrid
End Sub
Private Sub configurargrid()
    GridBusquedaMulta.Clear
    GridBusquedaMulta.FormatString = "cedula|nombre|apellido|fecha|nombre_multa"
    GridBusquedaMulta.ColWidth(0) = 1000
    GridBusquedaMulta.ColWidth(1) = 1000
    GridBusquedaMulta.ColWidth(2) = 1300
    GridBusquedaMulta.ColWidth(3) = 1200
    GridBusquedaMulta.ColWidth(4) = 1800
    'GridBusquedaCobranzas.ColWidth(5) = 1500

End Sub
Private Sub cargargrid()
    Dim sql As String
            sql = "select s.cedula,s.nombre,s.apellido,sm.fecha,m.nombre_multa from socio s join  socio_multas sm on s.idsocio=sm.idsocio join multas m on m.idmulta=sm.idmulta"
            Set TBLMulta = CONEXION.Execute(sql)
    Dim f As Integer
    f = 1
    GridBusquedaMulta.Rows = 2
    Do Until TBLMulta.EOF
        GridBusquedaMulta.TextMatrix(f, 0) = TBLMulta!cedula
        GridBusquedaMulta.TextMatrix(f, 1) = TBLMulta!nombre
        GridBusquedaMulta.TextMatrix(f, 2) = TBLMulta!apellido
        GridBusquedaMulta.TextMatrix(f, 3) = TBLMulta!fecha
        GridBusquedaMulta.TextMatrix(f, 4) = TBLMulta!nombre_multa
        'GridBusquedaCobranzas.TextMatrix(f, 5) = TBLCobrar!fecha
       
        TBLMulta.MoveNext
        f = f + 1
        GridBusquedaMulta.Rows = GridBusquedaMulta.Rows + 1
        
    Loop

End Sub
Private Sub GridBusquedaMulta_Click()
    Dim z As Integer
    z = GridBusquedaMulta.Row
    If z > 0 Then
        
        cimulta = GridBusquedaMulta.TextMatrix(GridBusquedaMulta.Row, 4)
        idmulta = ModuloFunciones.buscarID("socio_multas", "idsocio", cimulta)
        
        Set TBLMulta = Nothing
        TBLMulta.Open "select * from socio_multas where idsocio= '" & cimulta & "'", CONEXION, adOpenDynamic, adLockOptimistic
    

        TBLMulta.MoveFirst
        
        FrmCobranzasMultas.GridDetalle.Row = TBLMulta.Fields("nombre_multa").Value
        'FrmCobranzasMultas.txtnombre.Text = TBLCobrar.Fields("nombre").Value
        'FrmCobranzasMultas.txtapellido.Text = TBLCobrar.Fields("apellido").Value
        'FrmCobranzasMultas.txtValor.Text = TBLCobrar.Fields("valor").Value
        'FrmCobranzasMultas.txtsaldo.Text = TBLCobrar.Fields("saldo").Value
    Else
    End If
        Unload Me
    

End Sub

Private Sub txtBusqueda_Change()
    
    Dim sql As String
    Set TBLMulta = Nothing
    sql = "select s.cedula,s.nombre,s.apellido,sm.fecha,m.nombre_multa from  socio s join  socio_multas sm on s.idsocio=sm.idsocio join multas m on m.idmulta=sm.idmulta where cedula like '%" & txtBusqueda & "%' or nombre like '%" & Trim(UCase(txtBusqueda.Text)) & "%' or apellido like '%" & Trim(UCase(txtBusqueda.Text)) & "%'"
    Set TBLMulta = CONEXION.Execute(sql)
    
    Dim f As Integer
    f = 1
    GridBusquedaMulta.Rows = 2
        Do Until TBLMulta.EOF
            GridBusquedaMulta.TextMatrix(f, 0) = TBLMulta!cedula
            GridBusquedaMulta.TextMatrix(f, 1) = TBLMulta!nombre
            GridBusquedaMulta.TextMatrix(f, 2) = TBLMulta!apellido
             GridBusquedaMulta.TextMatrix(f, 3) = TBLMulta!fecha
             GridBusquedaMulta.TextMatrix(f, 4) = TBLMulta!nombre_multa
             'GridBusquedaMulta.TextMatrix(f, 5) = TBLCobrar!fecha
            TBLMulta.MoveNext
            f = f + 1
            GridBusquedaMulta.Rows = GridBusquedaMulta.Rows + 1
            
        Loop

End Sub


