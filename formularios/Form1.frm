VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGestionMultas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multas"
   ClientHeight    =   8520
   ClientLeft      =   1950
   ClientTop       =   1305
   ClientWidth     =   12435
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   12435
   ShowInTaskbar   =   0   'False
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
      Left            =   6240
      Picture         =   "Form1.frx":424A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6840
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker DTPanterior 
      Height          =   375
      Left            =   7200
      TabIndex        =   24
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111017985
      CurrentDate     =   43384
   End
   Begin MSComCtl2.DTPicker DTPactual 
      Height          =   375
      Left            =   10200
      TabIndex        =   22
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111017985
      CurrentDate     =   43384
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
      Left            =   4920
      TabIndex        =   19
      Top             =   3720
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox txtBUSCAR 
      Height          =   375
      Left            =   1680
      MaxLength       =   25
      TabIndex        =   18
      Top             =   4320
      Width           =   9735
   End
   Begin VB.TextBox txtapellido 
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtnombre 
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   1080
      Width           =   2775
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
      Left            =   3120
      TabIndex        =   13
      Top             =   3720
      Width           =   1335
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
      Left            =   1320
      TabIndex        =   12
      Top             =   3720
      Width           =   1095
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
      Left            =   4200
      Picture         =   "Form1.frx":4C15
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   1320
      TabIndex        =   7
      Top             =   6600
      Width           =   7455
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
         Picture         =   "Form1.frx":54DF
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   6120
         Picture         =   "Form1.frx":87E3
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "Form1.frx":91A9
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "Form1.frx":C403
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "Form1.frx":CE01
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1100
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridMultas 
      Height          =   1575
      Left            =   1320
      TabIndex        =   6
      Top             =   4920
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.ComboBox cmbMultas 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   2160
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2880
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111017985
      CurrentDate     =   43348
   End
   Begin VB.TextBox txtSocio 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   2775
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
      Left            =   240
      TabIndex        =   26
      Top             =   4320
      Width           =   1245
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
      Left            =   6120
      TabIndex        =   23
      Top             =   3720
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
      Left            =   8760
      TabIndex        =   21
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label Label5 
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
      Left            =   240
      TabIndex        =   17
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multa"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   525
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   525
   End
End
Attribute VB_Name = "frmGestionMultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tblsociomulta As New ADODB.Recordset
Dim idsocio_multas As Integer
Dim TBLMultas As New ADODB.Recordset
Dim TBLSocio As New ADODB.Recordset
Dim idmulta As Integer
Dim IDSOCIO As Integer
Dim cimulta As Integer
Dim TBLGuardarMultas As New ADODB.Recordset
Private Sub cmbMultas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DTPFecha.SetFocus
End Sub
Private Sub CmdAgregar1_Click()
    FrmAgregarMultas.Show
    frmGestionMultas.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    Call desactivarcajas
End Sub
Private Sub cmdCerrar_Click()
frmMenu.Show
frmGestionMultas.Hide
End Sub
Private Sub enviar_a_excel()
    Dim fso, f, i, objeto1
    Dim cadena As String
    Dim NUM As Integer
    Dim Objeto As Object
    
    Set Objeto = Nothing
    Set objeto1 = CreateObject("Excel.Application")
    objeto1.Visible = True
    objeto1.workbooks.Open FileName:=App.Path & "\Recibos\listadoDeudores.xls"
    cadena = App.Path & "\Recibos\listadoDeudores.xls"
    
     'objeto1.cells(4, 4).Value = Trim(Str(Date))
   ' Set Objeto = GetObject(cadena)
    'Objeto.Application.Windows("RECIBO.xlsx").Visible = True
    
    With objeto1 '.worksheets("RecibosMultas")
    
    
   'cells(8, 5).Value = Trim(Str(Date))
    '.cells(8, 2).Value = txtSocio.Text
    '.cells(8, 2).Value = Me.txtnombre.Text
    '.cells(8, 3).Value = Me.txtapellido.Text
    '.cells(8, 4).Value = Me.cmbMultas.Text
    
    
    
    '.cells(19, 4).Value = Me.lblValor.Caption
    '.cells(2, 5).Value = Me.LblNumeroRecibo.Caption
    
    
    NUM = 8
    Dim cod As Integer
    cod = 1
    For i = 1 To GridMultas.Rows - 1
        
        .cells(NUM, 1).Value = GridMultas.TextMatrix(i, 1)
        .cells(NUM, 2).Value = GridMultas.TextMatrix(i, 2)
        .cells(NUM, 3).Value = GridMultas.TextMatrix(i, 3)
        .cells(NUM, 4).Value = GridMultas.TextMatrix(i, 4)
        .cells(NUM, 5).Value = GridMultas.TextMatrix(i, 5)
        .cells(NUM, 6).Value = GridMultas.TextMatrix(i, 7)
        
        cod = cod + 1
        NUM = NUM + 1
        'End If
        'insertar fila en excel'
        'Objeto.Range(LTrim(Str(NUM)) & ":" & LTrim(Str(NUM + 1))).Select
        'Objeto.Selection.Insert
        
        
        
    
    Next
    
    
    End With
End Sub

Private Sub cmdImprimir_Click()
Call enviar_a_excel
End Sub

Private Sub cmdModificar_Click()
If Len(Trim(cmbMultas.Text)) > 0 Then
    
    Dim tblguardarsociomulta1 As New ADODB.Recordset
    Dim z As Integer
    z = MsgBox("¿Desea Guardar el Socio?", vbYesNo, "Gestion Multas")
    If z = vbYes Then
    'Set tblgurdarsociomulta1 = Nothing
    Dim respuestaID As Integer
    'Dim socio_ingresado As Integer
    'Dim multa_seleccionada As Integer
     'Dim observacion_ingresada As Integer
     'Dim valor_ingresado As Integer
     'observacion_ingresada = buscarID("multas", "observacion", Me.txtObservacion.Text)
     'valor_ingresado = buscarID("multas", "valor", Me.txtValor.Text)
     
     socio_multado = buscarID("socio_multas", "idsocio_multas", cimulta)
     multa_seleccionada = buscarID("multas", "nombre_multa", cmbMultas.Text)
        
        Set tblgurdarsociomulta1 = Nothing
        tblguardarsociomulta1.Open "select * from socio_multas where idsocio_multas=" & cimulta, CONEXION, adOpenDynamic, adLockOptimistic
        'tblguardarsociomulta1.AddNew
        tblguardarsociomulta1!idsocio_multas = socio_multado
        tblguardarsociomulta1!idmulta = multa_seleccionada
        'tblguardarsociomulta1!observacion = Me.txtObservacion.Text
        tblguardarsociomulta1!fecha = Me.DTPFecha.Value
       'tblguardarsociomulta1!valor = Me.txtValor.Text
       
       
       tblguardarsociomulta1.Update
       
       Call cargargrid
       Call desactivarcajas
       cmbMultas.Text = ""
    End If
  Else
    MsgBox "¡No se puede Guardar Datos vacios!,Ingresar valores en los campos", vbInformation, ""
  End If

End Sub

Private Sub cmdRegistro_Click()
    Call activarcajas
    Registro.Show
    'frmGestionMultas.Enabled = False
    
End Sub

Private Sub cmdReporte_Click()
    Dim nombreseccion As String
    Dim rsReporte As New ADODB.Recordset
    Dim sql As String
    
    Set rsReporte = Nothing
    If optTodos.Value = True Then
        sql = " select s.cedula,s.nombre,s.apellido,m.nombre_multa,sm.fecha,m.observacion,m.valor from socio s join socio_multas sm  on s.idsocio=sm.idsocio join  multas m on m.idmulta=sm.idmulta"
    End If
    If optcancelado.Value = True Then
        sql = "select sm.idsocio_multas,s.cedula,s.nombre,s.apellido,m.nombre_multa,sm.fecha,m.observacion,m.valor from socio s join socio_multas sm  on s.idsocio=sm.idsocio join  multas m on m.idmulta=sm.idmulta where  fecha<='" & Me.DTPactual.Value & "' and fecha >='" & Me.DTPanterior.Value & "' and estado_pago = 'CANCELADO' order by idsocio_multas asc"
    End If
    If optdebe.Value = True Then
        sql = "select sm.idsocio_multas,s.cedula,s.nombre,s.apellido,m.nombre_multa,sm.fecha,m.observacion,m.valor from socio s join socio_multas sm  on s.idsocio=sm.idsocio join  multas m on m.idmulta=sm.idmulta where  fecha<='" & Me.DTPactual.Value & "' and fecha >='" & Me.DTPanterior.Value & "' and estado_pago = 'DEBE' order by idsocio_multas asc"
    End If
     Set rsReporte = CONEXION.Execute(sql)
     If Not rsReporte Is Nothing Then
     
     nombreseccion = "sesion1"
     
     With ReporteMultas
     
     .LeftMargin = 0
     .RightMargin = 0
     .TopMargin = 0
     .BottomMargin = 0
     
     .Sections(nombreseccion).Controls.Item("txtcedula").DataField = "cedula"
     .Sections(nombreseccion).Controls.Item("txtnombre").DataField = "nombre"
     .Sections(nombreseccion).Controls.Item("txtapellido").DataField = "apellido"
     .Sections(nombreseccion).Controls.Item("txtmulta").DataField = "nombre_multa"
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
    Call CargarCombomultas
    Call desactivarcajas
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

Private Sub GridMultas_DblClick()
    Dim z As Integer
If optTodos.Value = True Then
        
 ElseIf optcancelado.Value = True Then MsgBox ("No se puede modificar el contenido")
End If
    If optdebe.Value = True Then
    z = MsgBox("¿Desea Modificar el dato registrado?", vbYesNo, "Gestion Multa")
    If z = vbYes Then
        cimulta = GridMultas.TextMatrix(GridMultas.Row, 0)
        idmulta = ModuloFunciones.buscarID("socio_multas", "idsocio_multas", cimulta)
        
            
        GridMultas.Col = 1
        txtSocio.Text = GridMultas.TextMatrix(GridMultas.RowSel, GridMultas.Col)
        GridMultas.Col = 2
        txtnombre.Text = GridMultas.TextMatrix(GridMultas.RowSel, GridMultas.Col)
        GridMultas.Col = 3
        txtapellido.Text = GridMultas.TextMatrix(GridMultas.RowSel, GridMultas.Col)
        GridMultas.Col = 4
        cmbMultas.Text = GridMultas.TextMatrix(GridMultas.RowSel, GridMultas.Col)
        'GridMultas.Col = 5
        'DTPFecha.Value = GridMultas.TextMatrix(GridMultas.RowSel, GridMultas.Col)
            
        Set TBLMultas = Nothing
        TBLMultas.Open " select * from socio_multas where idsocio_multas='" & cimulta & "'", CONEXION, adOpenDynamic, adLockOptimistic
        
        'If Not (TBLMultas.EOF) Then
            'TBLMultas.MoveFirst
        'Else
        
        GridMultas.Col = 1
        txtSocio.Text = GridMultas.TextMatrix(GridMultas.RowSel, GridMultas.Col)
        GridMultas.Col = 2
        txtnombre.Text = GridMultas.TextMatrix(GridMultas.RowSel, GridMultas.Col)
        GridMultas.Col = 3
        txtapellido.Text = GridMultas.TextMatrix(GridMultas.RowSel, GridMultas.Col)
        GridMultas.Col = 4
        cmbMultas.Text = GridMultas.TextMatrix(GridMultas.RowSel, GridMultas.Col)
        
        'txtSocio.Text = TBLMultas.Fields("idsocio").Value
        'txtnombre.Text = TBLMultas.Fields("idmulta").Value
        'txtapellido.Text=tblmultas.Filter(
        'txtObservacion.Text = TBLMultas.Fields("observacion").Value
        'Me.DTPFecha.Value = TBLMultas.Fields("fecha").Value
        
        'For i = 0 To Me.cmbMultas.ListCount - 1
            'cmbMultas.ListIndex = i
            'If Trim(UCase(cmbMultas.Text)) = UCase(ModuloFunciones.buscarID("multas", "idmulta", TBLMultas.Fields("idmulta").Value)) Then
                'Exit For
            'End If
        'Next
        
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
    'End If
End Sub

Private Sub Image1_Click()
    frmBusqueda.Show
    frmGestionMultas.Enabled = False
       
End Sub
Private Sub configurargrid()
    GridMultas.Clear
    GridMultas.FormatString = "Codigo|Cedula|Nombre|Apellido|Multa|Fecha|Observacion|valor"
    GridMultas.ColWidth(0) = 900
    GridMultas.ColWidth(1) = 1400
    GridMultas.ColWidth(2) = 1900
    GridMultas.ColWidth(3) = 2200
    GridMultas.ColWidth(4) = 2200
    GridMultas.ColWidth(5) = 1200
    GridMultas.ColWidth(6) = 2400
    GridMultas.ColWidth(7) = 800

End Sub
Private Sub cargargrid()
    Dim sql As String
        'sql = "select s.cedula,m.nombre_multa,observacion,sm.fecha,valor from socio s join socio_multas sm on sm.idsocio=sm.idsocio join multas m on m.idmulta=sm.idmulta"
            'sql = "select  socio.cedula,multas.nombre_multa,fecha,observacion from socio,socio_multas,multas where socio.idsocio=socio_multas.idsocio and multas.idmulta=socio_multas.idmulta"
            sql = " select sm.idsocio_multas,s.cedula,s.nombre,s.apellido,m.nombre_multa,sm.fecha,m.observacion,m.valor from socio s join socio_multas sm  on s.idsocio=sm.idsocio join  multas m on m.idmulta=sm.idmulta order by idsocio_multas asc"
            Set TBLMultas = CONEXION.Execute(sql)
    Dim f As Integer
    f = 1
    GridMultas.Rows = 2
    Do Until TBLMultas.EOF
        GridMultas.TextMatrix(f, 0) = TBLMultas!idsocio_multas
        GridMultas.TextMatrix(f, 1) = TBLMultas!cedula
        GridMultas.TextMatrix(f, 2) = TBLMultas!nombre
        GridMultas.TextMatrix(f, 3) = TBLMultas!apellido
        GridMultas.TextMatrix(f, 4) = TBLMultas!nombre_multa
        GridMultas.TextMatrix(f, 5) = IIf(IsNull(TBLMultas!fecha), "", TBLMultas!fecha)
        GridMultas.TextMatrix(f, 6) = IIf(IsNull(TBLMultas!observacion), "", TBLMultas!observacion)
        GridMultas.TextMatrix(f, 7) = TBLMultas!valor
        
       
        TBLMultas.MoveNext
        f = f + 1
        GridMultas.Rows = GridMultas.Rows + 1
        
    Loop

End Sub
Private Sub CargarCombomultas()
    cmbMultas.Clear
    Set TBLMultas = Nothing
    TBLMultas.Open "select * from multas", CONEXION, adOpenDynamic, adLockOptimistic
    Do Until TBLMultas.EOF
    cmbMultas.AddItem TBLMultas.Fields(1).Value
        TBLMultas.MoveNext
        Loop
End Sub
Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtValor.SetFocus
            KeyAscii = ModuloFunciones.letras(KeyAscii)
End Sub
Private Sub optcancelado_Click()
    Dim sql As String
    sql = "select sm.idsocio_multas,s.cedula,s.nombre,s.apellido,m.nombre_multa,sm.fecha,m.observacion,m.valor from socio s join socio_multas sm  on s.idsocio=sm.idsocio join  multas m on m.idmulta=sm.idmulta where  fecha<='" & Me.DTPactual.Value & "' and fecha >='" & Me.DTPanterior.Value & "' and estado_pago = 'CANCELADO' order by fecha"
        Set TBLSocio = CONEXION.Execute(sql)
        Dim f As Integer
        f = 1
        GridMultas.Rows = 2
        Do Until TBLSocio.EOF
            GridMultas.TextMatrix(f, 0) = TBLSocio!idsocio_multas
            GridMultas.TextMatrix(f, 1) = TBLSocio!cedula
            GridMultas.TextMatrix(f, 2) = TBLSocio!nombre
            GridMultas.TextMatrix(f, 3) = TBLSocio!apellido
            GridMultas.TextMatrix(f, 4) = TBLSocio!nombre_multa
            GridMultas.TextMatrix(f, 5) = TBLSocio!fecha
            GridMultas.TextMatrix(f, 6) = IIf(IsNull(TBLSocio!observacion), "", TBLSocio!observacion)
            GridMultas.TextMatrix(f, 7) = TBLSocio!valor
            TBLSocio.MoveNext
            f = f + 1
            GridMultas.Rows = GridMultas.Rows + 1
        Loop
 
End Sub
Private Sub optdebe_Click()
    Dim sql As String
    sql = "select sm.idsocio_multas,s.cedula,s.nombre,s.apellido,m.nombre_multa,sm.fecha,m.observacion,m.valor from socio s join socio_multas sm  on s.idsocio=sm.idsocio join  multas m on m.idmulta=sm.idmulta where  fecha<='" & Me.DTPactual.Value & "' and fecha >='" & Me.DTPanterior.Value & "' and estado_pago = 'DEBE' order by fecha"
    'sql = "select sm.idsocio_multas,s.cedula,s.nombre,s.apellido,m.nombre_multa,sm.fecha,m.observacion,m.valor from socio s join socio_multas sm  on s.idsocio=sm.idsocio join  multas m on m.idmulta=sm.idmulta where estado_pago = 'DEBE'"
        Set TBLSocio = CONEXION.Execute(sql)
        Dim f As Integer
        f = 1
        GridMultas.Rows = 2
        Do Until TBLSocio.EOF
            GridMultas.TextMatrix(f, 0) = TBLSocio!idsocio_multas
            GridMultas.TextMatrix(f, 1) = TBLSocio!cedula
            GridMultas.TextMatrix(f, 2) = TBLSocio!nombre
            GridMultas.TextMatrix(f, 3) = TBLSocio!apellido
            GridMultas.TextMatrix(f, 4) = TBLSocio!nombre_multa
            GridMultas.TextMatrix(f, 5) = TBLSocio!fecha
            GridMultas.TextMatrix(f, 6) = IIf(IsNull(TBLSocio!observacion), "", TBLSocio!observacion)
            GridMultas.TextMatrix(f, 7) = TBLSocio!valor
            TBLSocio.MoveNext
            f = f + 1
            GridMultas.Rows = GridMultas.Rows + 1
        Loop
        
End Sub
Private Sub optTodos_Click()
Dim sql As String
    sql = "select sm.idsocio_multas,s.cedula,s.nombre,s.apellido,m.nombre_multa,sm.fecha,m.observacion,m.valor from socio s join socio_multas sm  on s.idsocio=sm.idsocio join  multas m on m.idmulta=sm.idmulta  order by idsocio_multas asc"
        Set TBLSocio = CONEXION.Execute(sql)
        Dim f As Integer
        f = 1
        GridMultas.Rows = 2
        Do Until TBLSocio.EOF
            GridMultas.TextMatrix(f, 0) = TBLSocio!idsocio_multas
            GridMultas.TextMatrix(f, 1) = TBLSocio!cedula
            GridMultas.TextMatrix(f, 2) = TBLSocio!nombre
            GridMultas.TextMatrix(f, 3) = TBLSocio!apellido
            GridMultas.TextMatrix(f, 4) = TBLSocio!nombre_multa
            GridMultas.TextMatrix(f, 5) = IIf(IsNull(TBLSocio!fecha), "", TBLSocio!fecha)
            GridMultas.TextMatrix(f, 6) = IIf(IsNull(TBLSocio!observacion), "", TBLSocio!observacion)
            GridMultas.TextMatrix(f, 7) = TBLSocio!valor
            TBLSocio.MoveNext
            f = f + 1
            GridMultas.Rows = GridMultas.Rows + 1
        Loop
End Sub
Private Sub txtBuscar_Change()
    Dim sql As String
    Set TBLSocio = Nothing
    If Me.optTodos.Value = True Then
        sql = " select sm.idsocio_multas,s.cedula,s.nombre,s.apellido,m.nombre_multa,sm.fecha,m.observacion,m.valor from socio s join socio_multas sm  on s.idsocio=sm.idsocio join  multas m on m.idmulta=sm.idmulta where cedula like '%" & txtBuscar.Text & "%'  or  nombre like '%" & Trim(UCase(txtBuscar.Text)) & "%' or  apellido like '%" & Trim(UCase(txtBuscar.Text)) & "%'"
    End If
    If Me.optdebe.Value = True Then
        sql = " select sm.idsocio_multas,s.cedula,s.nombre,s.apellido,m.nombre_multa,sm.fecha,m.observacion,m.valor from socio s join socio_multas sm  on s.idsocio=sm.idsocio join  multas m on m.idmulta=sm.idmulta and estado_pago='DEBE' where cedula like '%" & txtBuscar.Text & "%'  or  nombre like '%" & Trim(UCase(txtBuscar.Text)) & "%' or  apellido like '%" & Trim(UCase(txtBuscar.Text)) & "%'"
        End If
    If Me.optcancelado.Value = True Then
        sql = "  select sm.idsocio_multas,s.cedula,s.nombre,s.apellido,m.nombre_multa,sm.fecha,m.observacion,m.valor from socio s join socio_multas sm  on s.idsocio=sm.idsocio join  multas m on m.idmulta=sm.idmulta and estado_pago='CANCELADO' where cedula like '%" & txtBuscar.Text & "%'  or  nombre like '%" & Trim(UCase(txtBuscar.Text)) & "%' or  apellido like '%" & Trim(UCase(txtBuscar.Text)) & "%'"
        End If
        
    
    
    Set TBLSocio = CONEXION.Execute(sql)
    
    Dim f As Integer
    f = 1
    GridMultas.Rows = 2
        Do Until TBLSocio.EOF
            GridMultas.TextMatrix(f, 0) = TBLSocio!idsocio_multas
            GridMultas.TextMatrix(f, 1) = TBLSocio!cedula
            GridMultas.TextMatrix(f, 2) = TBLSocio!nombre
            GridMultas.TextMatrix(f, 3) = TBLSocio!apellido
            GridMultas.TextMatrix(f, 4) = TBLSocio!nombre_multa
            GridMultas.TextMatrix(f, 5) = IIf(IsNull(TBLSocio!fecha), "", TBLSocio!fecha)
            GridMultas.TextMatrix(f, 6) = IIf(IsNull(TBLSocio!observacion), "", TBLSocio!observacion)
            GridMultas.TextMatrix(f, 7) = TBLSocio!valor
            TBLSocio.MoveNext
            f = f + 1
            GridMultas.Rows = GridMultas.Rows + 1
            
        Loop

End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
KeyAscii = ModuloFunciones.numeros_letras(KeyAscii)
End Sub

Private Sub txtSocio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmbMultas.SetFocus
            KeyAscii = ModuloFunciones.letras(KeyAscii)
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtBuscar.SetFocus
            KeyAscii = ModuloFunciones.Numeros(KeyAscii)
End Sub
Private Sub activarcajas()
    txtSocio.Enabled = False
    txtnombre.Enabled = False
    txtapellido.Enabled = False
    cmbMultas.Enabled = True
    DTPFecha.Enabled = True
'    txtBUSCAR.Enabled = False
    
    cmbMultas.Enabled = True
    'Image1.Enabled = True
    CmdAgregar1.Enabled = True
'    cmdGuardar.Enabled = True
    cmdModificar.Enabled = True
    cmdCancelar.Enabled = True
'    cmdAgregar.Enabled = False
End Sub
Private Sub desactivarcajas()
    txtSocio.Enabled = False
    cmbMultas.Enabled = False
    DTPFecha.Enabled = False
    txtnombre.Enabled = False
    txtapellido.Enabled = False
'    txtBUSCAR.Enabled = True
    
    
    txtSocio.Text = ""
    txtnombre.Text = ""
    txtapellido.Text = ""
    
    'Image1.Enabled = False
    Me.CmdAgregar1.Enabled = False
'    cmdGuardar.Enabled = False
    cmdModificar.Enabled = False
    cmdCancelar.Enabled = False
    cmdRegistro.Enabled = True
    cmdCancelar.Enabled = True
    cmdCerrar.Enabled = True
    cmdReporte.Enabled = True
    cmdImprimir.Enabled = True
'    cmdAgregar.Enabled = True
 'cmdAgregar.Enabled = True

End Sub
