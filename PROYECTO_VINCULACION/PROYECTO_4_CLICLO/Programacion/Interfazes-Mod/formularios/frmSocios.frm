VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Socios"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9540
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDireccion 
      Height          =   405
      Left            =   1920
      TabIndex        =   20
      Top             =   2760
      Width           =   4455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8160
      Top             =   6600
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PostgreSQL37"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PostgreSQL37"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Txtbuscar 
      Height          =   495
      Left            =   720
      TabIndex        =   18
      Top             =   5520
      Width           =   5175
   End
   Begin MSFlexGridLib.MSFlexGrid GridSocios 
      Height          =   1575
      Left            =   720
      TabIndex        =   17
      Top             =   6120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2778
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
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
      Height          =   735
      Left            =   7920
      Picture         =   "frmSocios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4440
      Width           =   1095
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
      Left            =   5640
      Picture         =   "frmSocios.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4320
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
      Left            =   3960
      Picture         =   "frmSocios.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
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
      Left            =   840
      Picture         =   "frmSocios.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   720
      TabIndex        =   11
      Top             =   4080
      Width           =   6615
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
         Left            =   1800
         Picture         =   "frmSocios.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtTelefono 
      Height          =   405
      Left            =   1920
      TabIndex        =   10
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtApellido 
      Height          =   405
      Left            =   1920
      TabIndex        =   9
      Top             =   2280
      Width           =   4455
   End
   Begin VB.TextBox txtNombre 
      Height          =   405
      Left            =   1920
      TabIndex        =   8
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtMovil 
      Height          =   405
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtCedula 
      Height          =   405
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
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
      Left            =   720
      TabIndex        =   19
      Top             =   2880
      Width           =   960
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TELEFONO"
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
      Left            =   720
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "APELLIDOS"
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
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRES"
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
      Left            =   720
      TabIndex        =   4
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIL"
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
      Left            =   945
      TabIndex        =   3
      Top             =   1440
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CEDULA"
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
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "REGISTRO DE SOCIO"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2925
      TabIndex        =   0
      Top             =   240
      Width           =   3045
   End
End
Attribute VB_Name = "frmSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idsocio As Integer
Dim cisocio As String
Dim TBLGuardarSocio As New ADODB.Recordset
Dim TBLSocio As New ADODB.Recordset
Private Sub cmdAgregar_Click()
    Call ActivarCajas
End Sub

Private Sub cmdCancelar_Click()
    Call DesactivarCajas
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
If Len(Trim(txtCedula.Text)) > 9 And Len(Trim(txtMovil.Text)) > 0 And Len(Trim(txtNombre.Text)) > 0 And Len(Trim(txtApellido.Text)) > 0 And Len(Trim(txtDireccion.Text)) > 0 And Len(Trim(txtDireccion)) > 0 Then
    
    Dim respuestaID As Integer
    Set TBLGuardarSocio = Nothing
    TBLGuardarSocio.Open "select * from socio where cedula = '" & txtCedula.Text & "'", CONEXION, adOpenDynamic, adLockOptimistic
    If Not (TBLGuardarSocio.EOF) Then
        MsgBox "Socio ya existe"
    Else
    Set TBLGuardarSocio = Nothing
    TBLGuardarSocio.Open "select * from socio", CONEXION, adOpenDynamic, adLockOptimistic
    TBLGuardarSocio.AddNew
    TBLGuardarSocio!cedula = txtCedula.Text
    TBLGuardarSocio!movil = txtMovil.Text
    TBLGuardarSocio!nombre = Trim(UCase(txtNombre.Text))
    TBLGuardarSocio!apellido = Trim(UCase(txtApellido.Text))
    TBLGuardarSocio!Direccion = Trim(UCase(txtDireccion.Text))
    TBLGuardarSocio!telefono = Trim(UCase(txtTelefono.Text))
    
    TBLGuardarSocio.Update
    
    Call DesactivarCajas
    Call CargarGrid
    
    End If
 Else
    MsgBox "Imposible guardar!.No se permiten campos vacios, ingresar valores en los campos", vbExclamation
  End If
    
End Sub
Private Sub cmdModificar_Click()
    If Len(Trim(txtCedula.Text)) > 9 And Len(Trim(txtMovil.Text)) > 0 And Len(Trim(txtNombre.Text)) > 0 And Len(Trim(txtApellido.Text)) > 0 And Len(Trim(txtDireccion.Text)) > 0 And Len(Trim(txtDireccion)) > 0 Then
      Dim z  As Integer
      z = MsgBox("¿Desea Guardar Socio?", vbYesNo, "Gestion Socios")
      If z = vbYes Then
        Dim respuestaID As Integer
        Set TBLGuardarSocio = Nothing
        
        Set TBLGuardarSocio = Nothing
        
        TBLGuardarSocio.Open "select * from socio where idsocio=" & idsocio, CONEXION, adOpenDynamic, adLockOptimistic

        TBLGuardarSocio!cedula = txtCedula.Text
        TBLGuardarSocio!movil = txtMovil.Text
        TBLGuardarSocio!nombre = Trim(UCase(txtNombre.Text))
        TBLGuardarSocio!apellido = Trim(UCase(txtApellido.Text))
        TBLGuardarSocio!Direccion = Trim(UCase(txtDireccion.Text))
        TBLGuardarSocio!telefono = Trim(UCase(txtTelefono.Text))
        
        TBLGuardarSocio.Update
        
        Call DesactivarCajas
        Call CargarGrid
        
    End If
 Else
    MsgBox "Imposible guardar!.No se permiten campos vacios, ingresar valores en los campos", vbExclamation
  End If
    
End Sub

Private Sub Form_Load()
    ModuloBaseDatos.conectardb
    Call DesactivarCajas
    Call ConfigurarGrid
    Call CargarGrid
End Sub
Private Sub GridSocios_DblClick()
    Dim z As Integer
    z = MsgBox("¿Desea Modificar el dato seleccionado?", vbYesNo, "Gestion Socios")
    If z = vbYes Then
        cisocio = GridSocios.TextMatrix(GridSocios.Row, 0)
        idsocio = ModuloFunciones.buscarID("socio", "cedula", cisocio)
        
        Set TBLSocio = Nothing
        TBLSocio.Open "select * from socio where cedula='" & cisocio & "'", CONEXION, adOpenDynamic, adLockOptimistic
        TBLSocio.MoveFirst
        
        txtCedula.Text = TBLSocio.Fields("cedula").Value
        txtMovil.Text = TBLSocio.Fields("movil").Value
        txtNombre.Text = IIf(IsNull(TBLSocio.Fields("nombre").Value), "", TBLSocio.Fields("nombre").Value)
        txtApellido.Text = IIf(IsNull(TBLSocio.Fields("apellido").Value), "", TBLSocio.Fields("apellido").Value)
        txtDireccion.Text = IIf(IsNull(TBLSocio.Fields("direccion").Value), "", TBLSocio.Fields("direccion").Value)
        txtTelefono.Text = IIf(IsNull(TBLSocio.Fields("telefono").Value), "", TBLSocio.Fields("telefono").Value)
    
         Call ActivarCajas
         cmdGuardar.Enabled = False
         cmdModificar.Enabled = True
         Me.cmdCancelar.Enabled = True
    Else
    
    
    End If

End Sub

Private Sub txtApellido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDireccion.SetFocus
        KeyAscii = ModuloFunciones.Direccion(KeyAscii)
End Sub
Private Sub Txtbuscar_Change()
    Dim sql As String
        sql = "select * from socio where cedula like '%" & Txtbuscar.Text & "'"
            Set TBLSocio = CONEXION.Execute(sql)
    Dim f As Integer
    f = 1
    GridSocios.Rows = 2
    Do Until TBLSocio.EOF
        GridSocios.TextMatrix(f, 0) = TBLSocio!cedula
        GridSocios.TextMatrix(f, 1) = TBLSocio!movil
        GridSocios.TextMatrix(f, 2) = TBLSocio!nombre
        GridSocios.TextMatrix(f, 3) = TBLSocio!apellido
        GridSocios.TextMatrix(f, 4) = TBLSocio!Direccion
        GridSocios.TextMatrix(f, 5) = TBLSocio!telefono
        TBLSocio.MoveNext
        f = f + 1
        GridSocios.Rows = GridSocios.Rows + 1
        
    Loop

End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtMovil.SetFocus
        KeyAscii = ModuloFunciones.Numeros(KeyAscii)
End Sub
Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTelefono.SetFocus
        KeyAscii = ModuloFunciones.Direccion(KeyAscii)
End Sub

Private Sub txtMovil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNombre.SetFocus
        KeyAscii = ModuloFunciones.Numeros(KeyAscii)
End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtApellido.SetFocus
        KeyAscii = ModuloFunciones.letras(KeyAscii)
End Sub
Private Sub ActivarCajas()
    txtCedula.Enabled = True
    txtNombre.Enabled = True
    Me.txtApellido.Enabled = True
    Me.txtDireccion.Enabled = True
    Me.txtTelefono.Enabled = True
    txtMovil.Enabled = True
   
    cmdGuardar.Enabled = True
    cmdModificar.Enabled = False
    cmdCancelar.Enabled = True
    cmdAgregar.Enabled = False
End Sub
Private Sub DesactivarCajas()

    txtCedula.Enabled = False
    txtNombre.Enabled = False
    txtMovil.Enabled = False
    Me.txtApellido.Enabled = False
    Me.txtDireccion.Enabled = False
    Me.txtTelefono.Enabled = False
    
    
    txtCedula.Text = ""
    txtNombre.Text = ""
    Me.txtApellido.Text = ""
    Me.txtDireccion.Text = ""
    Me.txtMovil.Text = ""
    Me.txtTelefono.Text = ""
    
    
    cmdGuardar.Enabled = False
    cmdModificar.Enabled = False
    cmdCancelar.Enabled = False
    cmdAgregar.Enabled = True
End Sub
Private Sub ConfigurarGrid()
    GridSocios.Clear
    GridSocios.FormatString = "cedula|movil|nombre|apellido|direccion|telefono"
    GridSocios.ColWidth(0) = 1000
    GridSocios.ColWidth(1) = 1500
    GridSocios.ColWidth(2) = 1500
    GridSocios.ColWidth(3) = 1500
    GridSocios.ColWidth(4) = 1500
    GridSocios.ColWidth(5) = 1500
    

End Sub
Private Sub CargarGrid()
    Dim sql As String
        sql = "select * from socio"
            Set TBLSocio = CONEXION.Execute(sql)
    Dim f As Integer
    f = 1
    GridSocios.Rows = 2
    Do Until TBLSocio.EOF
        GridSocios.TextMatrix(f, 0) = TBLSocio!cedula
        GridSocios.TextMatrix(f, 1) = TBLSocio!movil
        GridSocios.TextMatrix(f, 2) = TBLSocio!nombre
        GridSocios.TextMatrix(f, 3) = TBLSocio!apellido
        GridSocios.TextMatrix(f, 4) = TBLSocio!Direccion
        GridSocios.TextMatrix(f, 5) = TBLSocio!telefono
        TBLSocio.MoveNext
        f = f + 1
        GridSocios.Rows = GridSocios.Rows + 1
        
    Loop

End Sub

