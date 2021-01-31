VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAportaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aportaciones"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12045
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmAportaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBuscar 
      Height          =   375
      Left            =   1920
      MaxLength       =   25
      TabIndex        =   17
      Top             =   3840
      Width           =   6015
   End
   Begin VB.OptionButton optdebe 
      Caption         =   "Debe"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   3360
      Width           =   975
   End
   Begin VB.OptionButton optcancelado 
      Caption         =   "Cancelado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   3360
      Width           =   1335
   End
   Begin VB.OptionButton optTodos 
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   3360
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox txtnombre 
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtapellido 
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   1560
      Width           =   2535
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
      Height          =   495
      Left            =   4920
      Picture         =   "frmAportaciones.frx":424A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   480
      TabIndex        =   5
      Top             =   5880
      Width           =   7335
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Left            =   4920
         Picture         =   "frmAportaciones.frx":4B14
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "Reporte"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Left            =   3720
         Picture         =   "frmAportaciones.frx":54DF
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdRegistro 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Left            =   120
         Picture         =   "frmAportaciones.frx":8739
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Index           =   1
         Left            =   6120
         Picture         =   "frmAportaciones.frx":B922
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Left            =   1320
         Picture         =   "frmAportaciones.frx":C2E8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Left            =   2520
         Picture         =   "frmAportaciones.frx":CBEB
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1100
      End
   End
   Begin VB.ComboBox cmbAportaciones 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox txtSocio 
      Height          =   405
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid GridAportaciones 
      Height          =   1335
      Left            =   480
      TabIndex        =   9
      Top             =   4440
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   2640
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   112394241
      CurrentDate     =   43348
   End
   Begin MSComCtl2.DTPicker DTPanterior 
      Height          =   375
      Left            =   6000
      TabIndex        =   24
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   112394241
      CurrentDate     =   43384
   End
   Begin MSComCtl2.DTPicker DTPactual 
      Height          =   375
      Left            =   9120
      TabIndex        =   25
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   112394241
      CurrentDate     =   43384
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anterior"
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
      Left            =   4920
      TabIndex        =   27
      Top             =   3360
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Actual"
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
      Left            =   7680
      TabIndex        =   26
      Top             =   3240
      Width           =   1200
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
      Left            =   480
      TabIndex        =   21
      Top             =   3840
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Left            =   480
      TabIndex        =   13
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido"
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
      Left            =   480
      TabIndex        =   12
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Aportación"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Socio"
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
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   525
   End
End
Attribute VB_Name = "frmAportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLaportaciones As New ADODB.Recordset
Dim TBLSocio As New ADODB.Recordset
Dim idaportacion As Integer
Dim ciaportacion As Integer
Dim idsocio_aportacion As Integer
Dim TBLGuardarAportacion As New ADODB.Recordset
Dim tblsocioaportacion As New ADODB.Recordset
Private Sub cmbAportaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtHora.SetFocus
    KeyAscii = 0
End Sub
Private Sub cmdAgregar_Click()

End Sub
Private Sub CmdAgregar1_Click()
    FrmAgregarAportaciones.Show
    'frmAportaciones.Enabled = True
End Sub
Private Sub cmdCancelar_Click()
    Call desactivarcajas
End Sub
Private Sub cmdCerrar_Click(Index As Integer)
    frmMenu.Show
    frmAportaciones.Hide
End Sub

Private Sub cmdCerrar1_Click(Index As Integer)
    frmMenu.Show
    frmAportaciones.Hide
    
End Sub

Private Sub enviar_a_excel()
    Dim fso, f, i, objeto1
    Dim cadena As String
    Dim NUM As Integer
    Dim Objeto As Object
    
    Set Objeto = Nothing
    Set objeto1 = CreateObject("Excel.Application")
    objeto1.Visible = True
    objeto1.workbooks.Open FileName:=App.Path & "\Recibos\listadoDeudoresAportacion.xls"
    cadena = App.Path & "\Recibos\listadoDeudoresAportacion.xls"
    
    With objeto1 '.worksheets("RecibosMultas")
    
    NUM = 8
    Dim cod As Integer
    cod = 1
    For i = 1 To GridAportaciones.Rows - 1
        
        .cells(NUM, 1).Value = GridAportaciones.TextMatrix(i, 1)
        .cells(NUM, 2).Value = GridAportaciones.TextMatrix(i, 2)
        .cells(NUM, 3).Value = GridAportaciones.TextMatrix(i, 3)
        .cells(NUM, 4).Value = GridAportaciones.TextMatrix(i, 4)
        .cells(NUM, 5).Value = GridAportaciones.TextMatrix(i, 5)
        .cells(NUM, 6).Value = GridAportaciones.TextMatrix(i, 7)
        
        cod = cod + 1
        NUM = NUM + 1
    
    Next
     
    End With
End Sub
Private Sub Cmdguardar_Click()
'If Len(Me.cmbAportaciones.Text) > 1 And Me.DTPFecha.Value = False Then
'    Dim TBLguardarsocioaportacion As New ADODB.Recordset
'    socio_ingresado = buscarID("socio", "cedula", Me.txtSocio.Text)
'    aportacion_seleccionada = buscarID("aportaciones", "aportacion", Me.cmbAportaciones.Text)
'        Set TBLguardarsocioaportacion = Nothing
'        TBLguardarsocioaportacion.Open "select * from socio_aportaciones", CONEXION, adOpenDynamic, adLockOptimistic
'        TBLguardarsocioaportacion.AddNew
'        TBLguardarsocioaportacion!IDSOCIO = socio_ingresado
'        TBLguardarsocioaportacion!idaportaciones = aportacion_seleccionada
'        TBLguardarsocioaportacion!fecha = Me.DTPFecha.Value
'        TBLguardarsocioaportacion!estado_pago = "DEBE"
        'TBLguardarsocioaportacion!saldo = Me.
'        TBLguardarsocioaportacion.Update
'        TBLguardarsocioaportacion.Requery
        
'        Set TBLguardarsocioaportacion = Nothing
'        TBLguardarsocioaportacion.Open "select * from socio_aportaciones", CONEXION, adOpenDynamic, adLockOptimistic
'        TBLguardarsocioaportacion.MoveLast
'        idsocio_aportaciones = TBLguardarsocioaportacion!idsocio_aportaciones
'        Set TBLguardarsocioaportacion = Nothing
        
'        Set TBLAportacion = Nothing
'            sql = "update socio_aportaciones set estado_pago='DEBE' where idsocio_aportaciones='" & GridAportaciones.TextMatrix(GridAportaciones.Row, 0) & "'"
'        Set TBLAportacion = CONEXION.Execute(sql)
        
'        Call desactivarcajas
'        Call cargargrid
    'Else
        'MsgBox "Imposible guardar!.No se permiten campos vacios, ingresar valores en los campos", vbExclamation
  'End If
End Sub

Private Sub cmdImprimir_Click()
Call enviar_a_excel
End Sub

Private Sub cmdModificar_Click()
    If Len(Me.cmbAportaciones.Text) > 0 Then
        Dim tblguardarsocioaportaciones As New ADODB.Recordset
        Dim z As Integer
        z = MsgBox("¿Desea Guardar el Socio?", vbYesNo, "Aportaciones")
        If z = vbYes Then
            socio_ingresado = buscarID("socio_aportaciones", "idsocio_aportaciones", ciaportacion)
            aportacion_seleccionada = buscarID("aportaciones", "aportacion", Me.cmbAportaciones.Text)
            
            Set tblguardarsocioaportaciones = Nothing
            tblguardarsocioaportaciones.Open "select * from socio_aportaciones where idsocio_aportaciones=" & ciaportacion, CONEXION, adOpenDynamic, adLockOptimistic
            'tblguardarsocioaportaciones.AddNew
            tblguardarsocioaportaciones!idsocio_aportaciones = socio_ingresado
            tblguardarsocioaportaciones!idaportaciones = aportacion_seleccionada
            tblguardarsocioaportaciones!fecha = Me.DTPFecha.Value
            tblguardarsocioaportaciones.Update
            
            Call desactivarcajas
            Call cargargrid
            Me.cmbAportaciones.Text = ""
        End If
    Else
        MsgBox "¡No se puede Guardar Datos vacios!, Ingresar valores en los campos", vbInformation, ""
    End If
    
End Sub

Private Sub cmdRegistro_Click()
    Call activarcajas
    'CmbAportaciones.Enabled = True
    'Me.DTPFecha.Enabled = True
    'Me.CmdAgregar1.Enabled = True
    Frmregistroaportaciones.Show
    'frmAportaciones.Enabled = False
    
End Sub

Private Sub cmdReporte_Click()
    Dim nombreseccion As String
    Dim rsReporte As New ADODB.Recordset
    Dim sql As String
    
    Set rsReporte = Nothing
    If optTodos.Value = True Then
        sql = " select s.cedula,s.nombre,s.apellido,a.aportacion,sa.fecha,a.descripcion,a.valor from socio s join socio_aportaciones sa on s.idsocio=sa.idsocio join aportaciones a on a.idaportaciones=sa.idaportaciones"
    End If
    If optcancelado.Value = True Then
        sql = "select sa.idsocio_aportaciones ,s.cedula,s.nombre,s.apellido,a.aportacion,sa.fecha,a.descripcion,a.valor from socio s join socio_aportaciones sa on s.idsocio=sa.idsocio join aportaciones a on a.idaportaciones=sa.idaportaciones where fecha<='" & Me.DTPactual.Value & "' and fecha>='" & Me.DTPanterior.Value & "' and estado_pago= 'CANCELADO' order by idsocio_aportaciones asc"
    End If
    If optdebe.Value = True Then
        sql = "select sa.idsocio_aportaciones ,s.cedula,s.nombre,s.apellido,a.aportacion,sa.fecha,a.descripcion,a.valor from socio s join socio_aportaciones sa on s.idsocio=sa.idsocio join aportaciones a on a.idaportaciones=sa.idaportaciones where fecha<='" & Me.DTPactual.Value & "' and fecha>='" & Me.DTPanterior.Value & "' and estado_pago= 'DEBE' order by idsocio_aportaciones asc"
    End If
     Set rsReporte = CONEXION.Execute(sql)
     If Not rsReporte Is Nothing Then
     
     nombreseccion = "sesion1"
     
     With ReporteAportaciones
     
     .LeftMargin = 0
     .RightMargin = 0
     .TopMargin = 0
     .BottomMargin = 0
     
     .Sections(nombreseccion).Controls.Item("txtcedula").DataField = "cedula"
     .Sections(nombreseccion).Controls.Item("txtnombre").DataField = "nombre"
     .Sections(nombreseccion).Controls.Item("txtapellido").DataField = "apellido"
     .Sections(nombreseccion).Controls.Item("txtaportacion").DataField = "aportacion"
     .Sections(nombreseccion).Controls.Item("txtfecha").DataField = "fecha"
     Set .DataSource = rsReporte
     .Show vbModal
     
     Set rsReporte = Nothing
     
     End With
 End If

End Sub

Private Sub Form_Load()
    ModuloBaseDatos.conectardb
    Call configurargrid
    Call cargargrid
    Call desactivarcajas
    Call CargarComboaportaciones
  Me.DTPactual.Value = Date
    Me.DTPactual.MaxDate = Date
    Me.DTPactual.MinDate = Date - 365
    Me.DTPanterior.Value = Date
    Me.DTPanterior.MaxDate = Date
    Me.DTPanterior.MinDate = Date - 365
    DTPFecha.MaxDate = Date
    DTPFecha.MinDate = Date

    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMenu.Show
    Unload Me
End Sub

Private Sub GridAportaciones_DblClick()
   If optTodos.Value = True Then
        
 ElseIf optcancelado.Value = True Then MsgBox ("No se puede modificar el contenido")
End If
    If optdebe.Value = True Then
    Dim z As Integer
    z = MsgBox("¿Desea Modificar el dato registrado?", vbYesNo, "Gestion Aportaciones")
    If z = vbYes Then
        'GridAportaciones.Col = 0
        'txtSocio.Text = GridAportaciones.TextMatrix(GridAportaciones.RowSel, GridAportaciones.Col)
        ciaportacion = Me.GridAportaciones.TextMatrix(Me.GridAportaciones.Row, 0)
        idaportacion = ModuloFunciones.buscarID("socio_aportaciones", "idsocio_aportaciones", ciaportacion)
        GridAportaciones.Col = 1
        txtSocio.Text = GridAportaciones.TextMatrix(GridAportaciones.RowSel, GridAportaciones.Col)
        GridAportaciones.Col = 2
        txtnombre.Text = GridAportaciones.TextMatrix(GridAportaciones.RowSel, GridAportaciones.Col)
        GridAportaciones.Col = 3
        txtapellido.Text = GridAportaciones.TextMatrix(GridAportaciones.RowSel, GridAportaciones.Col)
        GridAportaciones.Col = 4
        cmbAportaciones.Text = GridAportaciones.TextMatrix(GridAportaciones.RowSel, GridAportaciones.Col)
        'Me.GridAportaciones.Col = 5
        'Me.DTPFecha.Value = GridAportaciones.TextMatrix(GridAportaciones.RowSel, GridAportaciones.Col)
        Set TBLaportaciones = Nothing
        TBLaportaciones.Open "select * from socio_aportaciones where idsocio_aportaciones='" & ciaportacion & "'", CONEXION, adOpenDynamic, adLockOptimistic
        GridAportaciones.Col = 1
        txtSocio.Text = GridAportaciones.TextMatrix(GridAportaciones.RowSel, GridAportaciones.Col)
        GridAportaciones.Col = 2
        txtnombre.Text = GridAportaciones.TextMatrix(GridAportaciones.RowSel, GridAportaciones.Col)
        GridAportaciones.Col = 3
        txtapellido.Text = GridAportaciones.TextMatrix(GridAportaciones.RowSel, GridAportaciones.Col)
        GridAportaciones.Col = 4
        cmbAportaciones.Text = GridAportaciones.TextMatrix(GridAportaciones.RowSel, GridAportaciones.Col)
        'cmdGuardar.Enabled = False
        
        Call activarcajas
        'Me.Cmdguardar.Enabled = False
        cmdRegistro.Enabled = False
        cmdImprimir.Enabled = False
        cmdReporte.Enabled = False
        cmdModificar.Enabled = True
        cmdCancelar.Enabled = True
    'Else
    
    End If
    End If
    
End Sub
Private Sub Image1_Click()
    frmBusquedaSocio.Show
    frmAportaciones.Enabled = False
End Sub
Private Sub CargarComboaportaciones()
    cmbAportaciones.Clear
    Set TBLaportaciones = Nothing
    TBLaportaciones.Open "select * from aportaciones", CONEXION, adOpenDynamic, adLockOptimistic
    Do Until TBLaportaciones.EOF
        cmbAportaciones.AddItem TBLaportaciones.Fields(2).Value
        TBLaportaciones.MoveNext
    Loop
End Sub
Private Sub optcancelado_Click()
    Dim sql As String
    sql = "select sa.idsocio_aportaciones ,s.cedula,s.nombre,s.apellido,a.aportacion,sa.fecha,a.descripcion,a.valor from socio s join socio_aportaciones sa on s.idsocio=sa.idsocio join aportaciones a on a.idaportaciones=sa.idaportaciones where fecha<='" & Me.DTPactual.Value & "' and fecha>='" & Me.DTPanterior.Value & "' and estado_pago = 'CANCELADO' order by fecha"
        Set TBLaportaciones = CONEXION.Execute(sql)
        Dim f As Integer
        f = 1
        GridAportaciones.Rows = 2
        Do Until TBLaportaciones.EOF
            GridAportaciones.TextMatrix(f, 0) = TBLaportaciones!idsocio_aportaciones
            GridAportaciones.TextMatrix(f, 1) = TBLaportaciones!cedula
            GridAportaciones.TextMatrix(f, 2) = TBLaportaciones!nombre
            GridAportaciones.TextMatrix(f, 3) = TBLaportaciones!apellido
            GridAportaciones.TextMatrix(f, 4) = TBLaportaciones!aportacion
            GridAportaciones.TextMatrix(f, 5) = TBLaportaciones!fecha
            GridAportaciones.TextMatrix(f, 6) = IIf(IsNull(TBLaportaciones!descripcion), "", TBLaportaciones!descripcion)
            GridAportaciones.TextMatrix(f, 7) = TBLaportaciones!valor
            TBLaportaciones.MoveNext
            f = f + 1
            GridAportaciones.Rows = GridAportaciones.Rows + 1
        Loop

 
End Sub
Private Sub optdebe_Click()
    Dim sql As String
    sql = "select sa.idsocio_aportaciones ,s.cedula,s.nombre,s.apellido,a.aportacion,sa.fecha,a.descripcion,a.valor from socio s join socio_aportaciones sa on s.idsocio=sa.idsocio join aportaciones a on a.idaportaciones=sa.idaportaciones where fecha<='" & Me.DTPactual.Value & "' and fecha>='" & Me.DTPanterior.Value & "' and estado_pago= 'DEBE' order by fecha"
        Set TBLaportaciones = CONEXION.Execute(sql)
        Dim f As Integer
        f = 1
        GridAportaciones.Rows = 2
        Do Until TBLaportaciones.EOF
            GridAportaciones.TextMatrix(f, 0) = TBLaportaciones!idsocio_aportaciones
            GridAportaciones.TextMatrix(f, 1) = TBLaportaciones!cedula
            GridAportaciones.TextMatrix(f, 2) = TBLaportaciones!nombre
            GridAportaciones.TextMatrix(f, 3) = TBLaportaciones!apellido
            GridAportaciones.TextMatrix(f, 4) = TBLaportaciones!aportacion
            GridAportaciones.TextMatrix(f, 5) = TBLaportaciones!fecha
            GridAportaciones.TextMatrix(f, 6) = IIf(IsNull(TBLaportaciones!descripcion), "", TBLaportaciones!descripcion)
            GridAportaciones.TextMatrix(f, 7) = TBLaportaciones!valor
            TBLaportaciones.MoveNext
            f = f + 1
            GridAportaciones.Rows = GridAportaciones.Rows + 1
        Loop
 
End Sub

Private Sub optTodos_Click()
    Dim sql As String
        sql = "select sa.idsocio_aportaciones ,s.cedula,s.nombre,s.apellido,a.aportacion,sa.fecha,a.descripcion,a.valor from socio s join socio_aportaciones sa on s.idsocio=sa.idsocio join aportaciones a on a.idaportaciones=sa.idaportaciones order by idsocio_aportaciones"
        Set TBLaportaciones = CONEXION.Execute(sql)
    Dim f As Integer
    f = 1
    GridAportaciones.Rows = 2
    Do Until TBLaportaciones.EOF
        GridAportaciones.TextMatrix(f, 0) = TBLaportaciones!idsocio_aportaciones
        GridAportaciones.TextMatrix(f, 1) = TBLaportaciones!cedula
        GridAportaciones.TextMatrix(f, 2) = TBLaportaciones!nombre
        GridAportaciones.TextMatrix(f, 3) = TBLaportaciones!apellido
        GridAportaciones.TextMatrix(f, 4) = IIf(IsNull(TBLaportaciones!aportacion), "", TBLaportaciones!aportacion)
        GridAportaciones.TextMatrix(f, 5) = TBLaportaciones!fecha
        GridAportaciones.TextMatrix(f, 6) = IIf(IsNull(TBLaportaciones!descripcion), "", TBLaportaciones!descripcion)
        GridAportaciones.TextMatrix(f, 7) = TBLaportaciones!valor
        
        TBLaportaciones.MoveNext
        f = f + 1
        GridAportaciones.Rows = GridAportaciones.Rows + 1
                
    Loop
End Sub

Private Sub txtBuscar_Change()
    Dim sql As String
    Set TBLSocio = Nothing
    If Me.optTodos.Value = True Then
        sql = "select sa.idsocio_aportaciones ,s.cedula,s.nombre,s.apellido,a.aportacion,sa.fecha,a.descripcion,a.valor from socio s join socio_aportaciones sa on s.idsocio=sa.idsocio join aportaciones a on a.idaportaciones=sa.idaportaciones where cedula like '%" & txtBuscar.Text & "%'  or  nombre like '%" & Trim(UCase(txtBuscar.Text)) & "%' or  apellido like '%" & Trim(UCase(txtBuscar.Text)) & "%'"
    End If
    If Me.optdebe.Value = True Then
        sql = "select sa.idsocio_aportaciones ,s.cedula,s.nombre,s.apellido,a.aportacion,sa.fecha,a.descripcion,a.valor from socio s join socio_aportaciones sa on s.idsocio=sa.idsocio join aportaciones a on a.idaportaciones=sa.idaportaciones  and estado_pago= 'DEBE' where cedula like '%" & txtBuscar.Text & "%'  or  nombre like '%" & Trim(UCase(txtBuscar.Text)) & "%' or  apellido like '%" & Trim(UCase(txtBuscar.Text)) & "%'"
    End If
    If Me.optcancelado.Value = True Then
        sql = "select sa.idsocio_aportaciones ,s.cedula,s.nombre,s.apellido,a.aportacion,sa.fecha,a.descripcion,a.valor from socio s join socio_aportaciones sa on s.idsocio=sa.idsocio join aportaciones a on a.idaportaciones=sa.idaportaciones  and estado_pago= 'CANCELADO' where cedula like '%" & txtBuscar.Text & "%'  or  nombre like '%" & Trim(UCase(txtBuscar.Text)) & "%' or  apellido like '%" & Trim(UCase(txtBuscar.Text)) & "%'"
    End If
    
    Set TBLSocio = CONEXION.Execute(sql)
    
    Dim f As Integer
    f = 1
    GridAportaciones.Rows = 2
        Do Until TBLSocio.EOF
            GridAportaciones.TextMatrix(f, 0) = TBLSocio!idsocio_aportaciones
            GridAportaciones.TextMatrix(f, 1) = TBLSocio!cedula
            GridAportaciones.TextMatrix(f, 2) = TBLSocio!nombre
            GridAportaciones.TextMatrix(f, 3) = TBLSocio!apellido
            GridAportaciones.TextMatrix(f, 4) = TBLSocio!aportacion
            GridAportaciones.TextMatrix(f, 5) = IIf(IsNull(TBLSocio!fecha), "", TBLSocio!fecha)
            GridAportaciones.TextMatrix(f, 6) = IIf(IsNull(TBLSocio!descripcion), "", TBLSocio!descripcion)
            GridAportaciones.TextMatrix(f, 7) = IIf(IsNull(TBLSocio!valor), "", TBLSocio!valor)
            TBLSocio.MoveNext
            f = f + 1
            GridAportaciones.Rows = GridAportaciones.Rows + 1
            
        Loop

End Sub
Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    KeyAscii = ModuloFunciones.numeros_letras(KeyAscii)
End Sub

Private Sub txtSocio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbAportaciones.SetFocus
    KeyAscii = ModuloFunciones.letras(KeyAscii)
End Sub
Private Sub desactivarcajas()

    Me.txtSocio.Enabled = False
    txtnombre.Enabled = False
    txtapellido.Enabled = False
    Me.cmbAportaciones.Enabled = False
    Me.DTPFecha.Enabled = False
    txtBuscar.Enabled = True
    
    CmdAgregar1.Enabled = False
    
    'Image1.Enabled = False
    'Me.CmdAgregar1.Enabled = False
'    cmdGuardar.Enabled = False
    cmdModificar.Enabled = False
    cmdCancelar.Enabled = False
    cmdRegistro.Enabled = True
    cmdCancelar.Enabled = True
    cmdReporte.Enabled = True
    cmdImprimir.Enabled = True
'    cmdAgregar.Enabled = True
 'cmdAgregar.Enabled = True

    
    
    
    Me.txtSocio.Text = ""
    txtnombre.Text = ""
    txtapellido.Text = ""
    'Me.GridAportaciones.Enabled = True
    
    
    
    '''''
    
    
    
    
    
    
    
End Sub
Private Sub activarcajas()

    Me.txtSocio.Enabled = False
    txtnombre.Enabled = False
    txtapellido.Enabled = False
    txtBuscar.Enabled = True
    Me.cmbAportaciones.Enabled = True
    Me.DTPFecha.Enabled = True
    CmdAgregar1.Enabled = True
    
    'Me.GridAportaciones.Enabled = True
    
    'cmdGuardar.Enabled = True
    cmdModificar.Enabled = True
    cmdCancelar.Enabled = True
    'CmdAgregar1.Enabled = True

End Sub
Private Sub configurargrid()
    GridAportaciones.Clear
    GridAportaciones.FormatString = "Codigo|Cedula|Nombre|Apellido|Aportacion|Fecha|Descripcion|Valor"
    GridAportaciones.ColWidth(0) = 700
    GridAportaciones.ColWidth(1) = 1200
    GridAportaciones.ColWidth(2) = 1700
    GridAportaciones.ColWidth(3) = 1700
    GridAportaciones.ColWidth(4) = 2200
    GridAportaciones.ColWidth(5) = 1400
    GridAportaciones.ColWidth(6) = 2000
    GridAportaciones.ColWidth(7) = 800
     
End Sub
Private Sub cargargrid()
    Dim sql As String
        sql = "select sa.idsocio_aportaciones ,s.cedula,s.nombre,s.apellido,a.aportacion,sa.fecha,a.descripcion,a.valor from socio s join socio_aportaciones sa on s.idsocio=sa.idsocio join aportaciones a on a.idaportaciones=sa.idaportaciones order by idsocio_aportaciones asc"
        Set TBLaportaciones = CONEXION.Execute(sql)
    Dim f As Integer
    f = 1
    GridAportaciones.Rows = 2
    Do Until TBLaportaciones.EOF
        GridAportaciones.TextMatrix(f, 0) = TBLaportaciones!idsocio_aportaciones
        GridAportaciones.TextMatrix(f, 1) = TBLaportaciones!cedula
        GridAportaciones.TextMatrix(f, 2) = TBLaportaciones!nombre
        GridAportaciones.TextMatrix(f, 3) = TBLaportaciones!apellido
        GridAportaciones.TextMatrix(f, 4) = IIf(IsNull(TBLaportaciones!aportacion), "", TBLaportaciones!aportacion)
        GridAportaciones.TextMatrix(f, 5) = TBLaportaciones!fecha
        GridAportaciones.TextMatrix(f, 6) = IIf(IsNull(TBLaportaciones!descripcion), "", TBLaportaciones!descripcion)
        GridAportaciones.TextMatrix(f, 7) = TBLaportaciones!valor
        
        TBLaportaciones.MoveNext
        f = f + 1
        GridAportaciones.Rows = GridAportaciones.Rows + 1
                
    Loop
    
End Sub

