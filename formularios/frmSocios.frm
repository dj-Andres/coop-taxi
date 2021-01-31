VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSocios 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Socios"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9240
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmSocios.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDireccion 
      Height          =   405
      Left            =   2280
      MaxLength       =   35
      TabIndex        =   16
      Top             =   2760
      Width           =   3255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7920
      Top             =   4800
      Visible         =   0   'False
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
      Height          =   375
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   14
      Top             =   3960
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid GridSocios 
      Height          =   1575
      Left            =   720
      TabIndex        =   13
      Top             =   4560
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2778
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   360
      TabIndex        =   11
      Top             =   6240
      Width           =   8655
      Begin VB.CommandButton cmdCerrar1 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   6960
         Picture         =   "frmSocios.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1300
         Left            =   3480
         Picture         =   "frmSocios.frx":499A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1500
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1300
         Left            =   5160
         Picture         =   "frmSocios.frx":9090
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1500
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1300
         Left            =   120
         Picture         =   "frmSocios.frx":D510
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1500
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1300
         Left            =   1800
         Picture         =   "frmSocios.frx":116DD
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.TextBox txtTelefono 
      Height          =   405
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   10
      Top             =   3240
      Width           =   3255
   End
   Begin VB.TextBox txtApellido 
      Height          =   405
      Left            =   2280
      MaxLength       =   25
      TabIndex        =   9
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtNombre 
      Height          =   405
      Left            =   2280
      MaxLength       =   25
      TabIndex        =   8
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox txtMovil 
      Height          =   405
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtCedula 
      Height          =   405
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar Socio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   750
      TabIndex        =   21
      Top             =   4080
      Width           =   1245
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      TabIndex        =   15
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   735
      TabIndex        =   6
      Top             =   3360
      Width           =   825
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellidos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   735
      TabIndex        =   5
      Top             =   2400
      Width           =   885
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombres"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   735
      TabIndex        =   4
      Top             =   1920
      Width           =   885
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Móvil"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   705
      TabIndex        =   3
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   735
      TabIndex        =   2
      Top             =   960
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GESTION DE SOCIOS"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3015
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IDSOCIO As Integer
Dim cisocio As Integer
Dim TBLGuardarSocio As New ADODB.Recordset
Dim TBLSocio As New ADODB.Recordset
Private Sub cmdAgregar_Click()
    Call activarcajas
End Sub
Private Sub cmdCancelar_Click()
    Call desactivarcajas
End Sub

Private Sub cmdCerrar1_Click()
frmMenu.Show
frmSocios.Hide
End Sub
Private Sub Cmdguardar_Click()
If Len(Trim(txtCedula.Text)) > 9 And Len(Trim(txtMovil.Text)) > 0 And Len(Trim(txtNombre.Text)) > 0 And Len(Trim(txtApellido.Text)) > 0 And Len(Trim(txtDireccion.Text)) > 0 And Len(Trim(txtDireccion)) > 0 Then
    
    Dim respuestaID As Integer
    Set TBLGuardarSocio = Nothing
    TBLGuardarSocio.Open "select * from socio where cedula = '" & txtCedula.Text & "'", CONEXION, adOpenDynamic, adLockOptimistic
    If Not (TBLGuardarSocio.EOF) Then
        MsgBox "Socio ya existe"
    Else
    Set TBLGuardarSocio = Nothing
    TBLGuardarSocio.Open "select * from socio ", CONEXION, adOpenDynamic, adLockOptimistic
    TBLGuardarSocio.AddNew
    TBLGuardarSocio!cedula = txtCedula.Text
    TBLGuardarSocio!movil = txtMovil.Text
    TBLGuardarSocio!nombre = Trim(UCase(txtNombre.Text))
    TBLGuardarSocio!apellido = Trim(UCase(txtApellido.Text))
    TBLGuardarSocio!Direccion = Trim(UCase(txtDireccion.Text))
    TBLGuardarSocio!telefono = Trim(UCase(txtTelefono.Text))
    
    TBLGuardarSocio.Update
    
    Call desactivarcajas
    Call cargargrid
    
    End If
 Else
    MsgBox "Imposible guardar!.No se permiten campos vacios, ingresar valores en los campos", vbExclamation
  End If
    
End Sub

Private Sub cmdModificar_Click()
If Len(Trim(txtCedula.Text)) > 9 And Len(Trim(txtMovil.Text)) > 0 And Len(Trim(txtNombre.Text)) > 0 And Len(Trim(txtApellido.Text)) > 0 And Len(Trim(txtDireccion.Text)) > 0 And Len(Trim(txtDireccion)) > 0 Then
    
    Dim respuestaID As Integer
    'Set TBLGuardarSocio = Nothing
    'TBLGuardarSocio.Open "select * from socio where cedula = '" & txtCedula.Text & "'", CONEXION, adOpenDynamic, adLockOptimistic
    'If Not (TBLGuardarSocio.EOF) Then
       ' MsgBox "Socio ya existe"
    'Else
    Set TBLGuardarSocio = Nothing
    TBLGuardarSocio.Open "select * from socio where idsocio=" & IDSOCIO, CONEXION, adOpenDynamic, adLockOptimistic
    'TBLGuardarSocio.AddNew
    TBLGuardarSocio!cedula = txtCedula.Text
    TBLGuardarSocio!movil = txtMovil.Text
    TBLGuardarSocio!nombre = Trim(UCase(txtNombre.Text))
    TBLGuardarSocio!apellido = Trim(UCase(txtApellido.Text))
    TBLGuardarSocio!Direccion = Trim(UCase(txtDireccion.Text))
    TBLGuardarSocio!telefono = Trim(UCase(txtTelefono.Text))
    
    TBLGuardarSocio.Update
    
    Call desactivarcajas
    Call cargargrid
    
    'End If
 Else
    MsgBox "Imposible guardar!.No se permiten campos vacios, ingresar valores en los campos", vbExclamation
  End If
    
End Sub
Private Sub Form_Load()
    ModuloBaseDatos.conectardb
    Call desactivarcajas
    Call configurargrid
    Call cargargrid
End Sub

    Private Sub Form_Unload(Cancel As Integer)
    frmMenu.Show
    Unload Me
End Sub

Private Sub GridSocios_DblClick()
    Dim z As Integer
    z = MsgBox("¿Desea Modificar el dato seleccionado?", vbYesNo, "Gestion Socios")
    If z = vbYes Then
    
    
    
        'GridSocios.Col = 0
        'Me.txtcedula.Text = GridSocios.TextMatrix(GridSocios.RowSel, GridSocios.Col)
        'GridSocios.Col = 1
        'txtMovil.Text = GridSocios.TextMatrix(GridSocios.RowSel, GridSocios.Col)
        'GridSocios.Col = 2
        'txtnombre.Text = GridSocios.TextMatrix(GridSocios.RowSel, GridSocios.Col)
        'GridSocios.Col = 3
        'txtapellido.Text = GridSocios.TextMatrix(GridSocios.RowSel, GridSocios.Col)
        'GridSocios.Col = 4
        'txtDireccion.Text = GridSocios.TextMatrix(GridSocios.RowSel, GridSocios.Col)
        'GridSocios.Col = 5
        'txtTelefono.Text = GridSocios.TextMatrix(GridSocios.RowSel, GridSocios.Col)
        
        cisocio = GridSocios.TextMatrix(GridSocios.Row, 0)
        IDSOCIO = ModuloFunciones.buscarID("socio", "idsocio", cisocio)
        
        Set TBLSocio = Nothing
        TBLSocio.Open "select * from socio where idsocio='" & cisocio & "'", CONEXION, adOpenDynamic, adLockOptimistic
        
        TBLSocio.MoveFirst
        
         
        txtCedula.Text = TBLSocio.Fields("cedula").Value
        txtMovil.Text = TBLSocio.Fields("movil").Value
        txtNombre.Text = IIf(IsNull(TBLSocio.Fields("nombre").Value), "", TBLSocio.Fields("nombre").Value)
        txtApellido.Text = IIf(IsNull(TBLSocio.Fields("apellido").Value), "", TBLSocio.Fields("apellido").Value)
        txtDireccion.Text = IIf(IsNull(TBLSocio.Fields("direccion").Value), "", TBLSocio.Fields("direccion").Value)
        txtTelefono.Text = IIf(IsNull(TBLSocio.Fields("telefono").Value), "", TBLSocio.Fields("telefono").Value)
    
         Call activarcajas
         cmdGuardar.Enabled = False
         cmdModificar.Enabled = True
         Me.cmdCancelar.Enabled = True
    Else
        MsgBox "No se puede guardar datos vacios", vbInformation, "Gestion de Socios"
    
    End If
End Sub

Private Sub txtApellido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDireccion.SetFocus
        KeyAscii = ModuloFunciones.Direccion(KeyAscii)
End Sub

Private Sub txtBuscar_Change()
     Dim sql As String
    Set TBLSocio = Nothing
    sql = "select * from socio where cedula like '%" & Txtbuscar.Text & "%'  or  nombre like '%" & Trim(UCase(Txtbuscar.Text)) & "%' or  apellido like '%" & Trim(UCase(Txtbuscar.Text)) & "%'  "
    Set TBLSocio = CONEXION.Execute(sql)
    
    Dim f As Integer
    f = 1
    GridSocios.Rows = 2
        Do Until TBLSocio.EOF
            GridSocios.TextMatrix(f, 0) = TBLSocio!movil
            GridSocios.TextMatrix(f, 1) = TBLSocio!cedula
            GridSocios.TextMatrix(f, 2) = TBLSocio!nombre
            GridSocios.TextMatrix(f, 3) = TBLSocio!apellido
            TBLSocio.MoveNext
            f = f + 1
            GridSocios.Rows = GridSocios.Rows + 1
            
        Loop

End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
KeyAscii = ModuloFunciones.numeros_letras(KeyAscii)
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
Private Sub activarcajas()
    txtCedula.Enabled = True
    txtNombre.Enabled = True
    Me.txtApellido.Enabled = True
    Me.txtDireccion.Enabled = True
    Me.txtTelefono.Enabled = True
    txtMovil.Enabled = True
    Me.Txtbuscar.Enabled = False
   
    cmdGuardar.Enabled = True
    cmdModificar.Enabled = False
    cmdCancelar.Enabled = True
    cmdAgregar.Enabled = False
End Sub
Private Sub desactivarcajas()

    txtCedula.Enabled = False
    txtNombre.Enabled = False
    txtMovil.Enabled = False
    Me.txtApellido.Enabled = False
    Me.txtDireccion.Enabled = False
    Me.txtTelefono.Enabled = False
    Me.Txtbuscar.Enabled = True
    
    
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
Private Sub configurargrid()
    GridSocios.Clear
    GridSocios.FormatString = "Codigo|Cedula|Movil|Nombre|Apellido|Direccion|Telefono"
    GridSocios.ColWidth(0) = 800
    GridSocios.ColWidth(1) = 1000
    GridSocios.ColWidth(2) = 800
    GridSocios.ColWidth(3) = 1500
    GridSocios.ColWidth(4) = 1500
    GridSocios.ColWidth(5) = 1500
    GridSocios.ColWidth(6) = 1500
    

End Sub
Private Sub cargargrid()
Dim sql As String
    sql = "select * from socio"
        Set TBLSocio = CONEXION.Execute(sql)
Dim f As Integer
f = 1
GridSocios.Rows = 2
Do Until TBLSocio.EOF
    GridSocios.TextMatrix(f, 0) = TBLSocio!IDSOCIO
    GridSocios.TextMatrix(f, 1) = TBLSocio!cedula
    GridSocios.TextMatrix(f, 2) = TBLSocio!movil
    GridSocios.TextMatrix(f, 3) = TBLSocio!nombre
    GridSocios.TextMatrix(f, 4) = TBLSocio!apellido
    GridSocios.TextMatrix(f, 5) = TBLSocio!Direccion
    GridSocios.TextMatrix(f, 6) = TBLSocio!telefono
    TBLSocio.MoveNext
    f = f + 1
    GridSocios.Rows = GridSocios.Rows + 1
    
Loop

End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
KeyAscii = ModuloFunciones.Numeros(KeyAscii)
End Sub
