VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubgrid.ocx"
Begin VB.Form FrmCobranzasMultas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cobranzas  Multas"
   ClientHeight    =   9585
   ClientLeft      =   4335
   ClientTop       =   855
   ClientWidth     =   11130
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmCobranzasMultas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "FrmCobranzasMultas.frx":424A
   ScaleHeight     =   9585
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtObservacion 
      Height          =   615
      Left            =   2880
      MaxLength       =   25
      TabIndex        =   20
      Top             =   6000
      Width           =   5895
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   5400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   110100481
      CurrentDate     =   43372
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar Multas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox txtapellido 
      Height          =   405
      Left            =   2520
      TabIndex        =   12
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox txtnombre 
      Height          =   405
      Left            =   2520
      TabIndex        =   11
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   960
      TabIndex        =   5
      Top             =   7680
      Width           =   8415
      Begin VB.CommandButton cmdCerrar 
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
         Height          =   1300
         Left            =   6720
         Picture         =   "FrmCobranzasMultas.frx":458C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1500
      End
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
         Height          =   1300
         Left            =   5040
         Picture         =   "FrmCobranzasMultas.frx":865C
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Picture         =   "FrmCobranzasMultas.frx":C327
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Cobrar"
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
         Picture         =   "FrmCobranzasMultas.frx":101DD
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   3360
         Picture         =   "FrmCobranzasMultas.frx":13DFE
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdBuscar 
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
      Left            =   5040
      Picture         =   "FrmCobranzasMultas.frx":1827E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtcedula 
      Height          =   405
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin ubGridControl.ubGrid UBMULTAS 
      Height          =   1575
      Left            =   840
      TabIndex        =   24
      Top             =   3600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2778
      Rows            =   1
      Cols            =   5
      Redraw          =   -1  'True
      ShowGrid        =   -1  'True
      GridSolid       =   -1  'True
      GridLineColor   =   12632256
      BackColorFixed  =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontEdit {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblNumeroRecibo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   515
      Left            =   8160
      TabIndex        =   26
      Top             =   840
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N. Recibo"
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
      Left            =   6960
      TabIndex        =   25
      Top             =   960
      Width           =   945
   End
   Begin VB.Label LBLDEUDA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   515
      Left            =   7800
      TabIndex        =   14
      Top             =   6840
      Width           =   600
   End
   Begin VB.Label lblsaldo 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Pendiente"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   23
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observación"
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
      Left            =   960
      TabIndex        =   19
      Top             =   6240
      Width           =   1155
   End
   Begin VB.Label Label1 
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
      Left            =   960
      TabIndex        =   17
      Top             =   5520
      Width           =   555
   End
   Begin VB.Label lblValor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   515
      Left            =   2880
      TabIndex        =   16
      Top             =   6840
      Width           =   600
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multas"
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
      Left            =   960
      TabIndex        =   13
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label6 
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
      Left            =   960
      TabIndex        =   10
      Top             =   2280
      Width           =   765
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
      Left            =   960
      TabIndex        =   9
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
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
      Left            =   960
      TabIndex        =   3
      Top             =   6960
      Width           =   510
   End
   Begin VB.Label Label2 
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
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   645
   End
   Begin VB.Label lblCobranzasMultas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COBRANZAS DE MULTAS"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "FrmCobranzasMultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONTARREGISTRO As Integer
Dim TBLGuardarCobro As New ADODB.Recordset
Dim VALORPAGAR As Double
Dim contarfilas As Integer
Dim TBLCobranzas As New ADODB.Recordset
Dim deuda As Double
Dim idcobranza_multa1 As Integer
Dim tblsociomulta As New ADODB.Recordset
Dim numeroRecibo As Integer
Dim BAN_imprimir As String
Private Sub cmdAgregar_Click()
    Call activarcajas
    frmBuscarCobranzas.Show
    FrmCobranzasMultas.Enabled = False
    Call ObtenerNumeroRecibo
End Sub
Private Sub cmdBuscar_Click()
    frmBuscarCobranzas.Show
End Sub

Private Sub cmdCancelar_Click()
    Call desactivarcajas
End Sub

Private Sub cmdCerrar_Click()
frmMenu.Show




FrmCobranzasMultas.Hide
End Sub
Private Sub Cmdguardar_Click()
    
  'If Me.DTPFecha.Value = False And Len(Me.txtObservacion.Text) And lblValor.Caption = 0 Then
    
    
    'Dim tblguardarsociomulta1 As New ADODB.Recordset
    'Dim socio_ingresado As Integer
    'Dim multa_seleccionada As Integer
     'Dim observacion_ingresada As Integer
     'Dim valor_ingresado As Integer
     'observacion_ingresada = buscarID("multas", "observacion", Me.txtObservacion.Text)
     'valor_ingresado = buscarID("multas", "valor", Me.txtValor.Text)
     
     socio_multado = buscarID("socio", "idsocio", IDSOCIO1)
     'multa_seleccionada = buscarID("multas", "nombre_multa", Me.cmbMultas.Text)
        
        Set TBLGuardarCobro = Nothing
        TBLGuardarCobro.Open "select * from cobranzas_multas", CONEXION, adOpenDynamic, adLockOptimistic
        TBLGuardarCobro.AddNew
       TBLGuardarCobro!IDSOCIO = socio_multado
        TBLGuardarCobro!fecha = DTPFecha.Value
        TBLGuardarCobro!observacion = Me.txtObservacion.Text
        TBLGuardarCobro!valor = Me.lblValor.Caption
        TBLGuardarCobro!idusuario = 1 ' IDUSUARIO1
        TBLGuardarCobro!saldo = Me.LBLDEUDA.Caption
        TBLGuardarCobro!numero_recibo = Me.LblNumeroRecibo.Caption
        
         
        TBLGuardarCobro.Update
        TBLGuardarCobro.Requery
        
        Set TBLGuardarCobro = Nothing
        TBLGuardarCobro.Open "select * from cobranzas_multas", CONEXION, adOpenDynamic, adLockOptimistic
        TBLGuardarCobro.MoveLast
        idcobranza_multa1 = TBLGuardarCobro!idcobranzas_multas
        Set TBLGuardarCobro = Nothing
        
        
        
        Set tblsociomulta = Nothing
        
        
        Dim sql As String
           
        Dim a As Integer
        
        For a = 1 To UBMULTAS.Rows
        If UBMULTAS.TextMatrix(a, 1) = 1 Then
            sql = "update  socio_multas set estado_pago='CANCELADO' where idsocio_multas='" & UBMULTAS.TextMatrix(a, 6) & "'"
            Set tblsociomulta = CONEXION.Execute(sql)
               sql = "update  socio_multas set referencia='" & idcobranza_multa1 & "' where idsocio_multas='" & UBMULTAS.TextMatrix(a, 6) & "'"
            Set tblsociomulta = CONEXION.Execute(sql)
               
        End If
    Next
        MsgBox "EL PAGO HA SIDO GUARDADO EXITOSAMENTE", vbInformation, "COBROS DE MULTAS"
        
        'Call cargargridMULTAS
    cmdImprimir.Enabled = True
    Cmdguardar.Enabled = False
    UBMULTAS.Enabled = False
    Me.Command1.Enabled = False
    Me.DTPFecha.Enabled = False
    Me.txtObservacion.Enabled = False
    
    
    Call GenerarNUEVONumeroRecibo
    
End Sub
Private Sub cmdImprimir_Click()
    Call enviar_a_excel
    UBMULTAS.Clear
    txtcedula.Text = ""
    txtnombre.Text = ""
    txtapellido.Text = ""
    DTPFecha.Enabled = False
    Me.txtObservacion.Text = ""
    Me.lblValor.Caption = ""

End Sub
Private Sub Command1_Click()
    contarfilas = 1
    Call configurargridMULTAS
    Call cargargridMULTAS
    UBMULTAS.Enabled = True
End Sub
Private Sub DTPFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtObservacion.SetFocus
End Sub
Private Sub Form_Load()
    ModuloBaseDatos.conectardb
    Call desactivarcajas
    'Call configurargrid
    'Call cargargrid
    'Call configurargriddetalle
    contarfilas = 1
     DTPFecha.Value = Date
     
     Me.DTPFecha.MinDate = Date - 1
     Me.DTPFecha.MaxDate = Date
End Sub
Private Sub activarcajas()
    Me.txtcedula.Enabled = False
    txtnombre.Enabled = False
    txtapellido.Enabled = False
    Me.DTPFecha.Enabled = True
    txtObservacion.Enabled = True
    Me.Command1.Enabled = True
    cmdImprimir.Enabled = False
    'txtValor.Enabled = True
    'txtsaldo.Enabled = True
    
    
    
    Cmdguardar.Enabled = True
    cmdCancelar.Enabled = True
    cmdBuscar.Enabled = True
End Sub
Private Sub desactivarcajas()
    Me.txtcedula.Enabled = False
    Me.txtcedula.Enabled = False
    txtnombre.Enabled = False
    txtapellido.Enabled = False
    Me.DTPFecha.Enabled = False
    'Me.DTPFecha.Enabled = False
    txtObservacion.Enabled = False
    'txtValor.Enabled = False
    'txtsaldo.Enabled = False
    
    
    cmdBuscar.Enabled = False
    cmdImprimir.Enabled = False
    Me.Command1.Enabled = False
    txtcedula.Text = ""
    txtnombre.Text = ""
    txtapellido.Text = ""
    'DTPFecha.Value = ""
    'txtobservacion.Text = ""
    'txtValor.Text = ""
    'txtValor.Text = ""
    
        UBMULTAS.Enabled = False
    
    
    Cmdguardar.Enabled = False
    cmdCancelar.Enabled = False
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMenu.Show
FrmCobranzasMultas.Hide
FrmCobranzasMultas.Enabled = True

End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then lblValor.SetFocus
        KeyAscii = ModuloFunciones.letras(KeyAscii)
End Sub
Private Sub txtSocio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.DTPFecha.SetFocus
        KeyAscii = ModuloFunciones.Numeros(KeyAscii)
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TXTSALDO.SetFocus
        KeyAscii = ModuloFunciones.Numeros(KeyAscii)
End Sub
Private Sub configurargrid()
    GridCobranzas.Clear
    GridCobranzas.FormatString = "socio|observacion|fecha|valor|saldo"
    GridCobranzas.ColWidth(0) = 1500
    GridCobranzas.ColWidth(1) = 2500
    GridCobranzas.ColWidth(2) = 1500
    GridCobranzas.ColWidth(3) = 1000
    GridCobranzas.ColWidth(4) = 1500
End Sub
Private Sub configurargridMULTAS()
    UBMULTAS.AutoSetup 2, 7, True, True, "Marcar|fecha|nombre_multa|valor|observacion|Codigo"
    
        UBMULTAS.ColMask(1) = checkmark
        UBMULTAS.ColMask(4) = NumericOnly
        UBMULTAS.ColMask(6) = NumericOnly
        
    UBMULTAS.ColWidth(1) = 40
    UBMULTAS.ColWidth(2) = 80
    UBMULTAS.ColWidth(3) = 160
    UBMULTAS.ColWidth(4) = 70
    UBMULTAS.ColWidth(5) = 150
    UBMULTAS.ColWidth(6) = 90
    
    'UBMULTAS.AutoRedraw = False

    UBMULTAS.ColAllowEdit(2) = False
    UBMULTAS.ColAllowEdit(3) = False
    UBMULTAS.ColAllowEdit(4) = False
    UBMULTAS.ColAllowEdit(5) = False
    UBMULTAS.ColAllowEdit(6) = False
    
    
End Sub
Private Sub cargargridMULTAS()
    Dim sql As String
            sql = "select sm.fecha,m.nombre_multa,m.valor,m.observacion,sm.idsocio_multas,sm.estado_pago from socio_multas sm join multas m on m.idmulta=sm.idmulta where sm.estado_pago= 'DEBE' and idsocio='" & IDSOCIO1 & "'"
           
            Set TBLMulta = CONEXION.Execute(sql)
    Dim f As Integer
    If contarfilas > 1 Then
        With UBMULTAS
       ' .Rows = .Rows + 1 ' significa la cantidad de fila
       ' .Row = .Rows - 1 ' significa la fila
        End With
    End If
    
    f = 1
    UBMULTAS.Rows = 0
        Do Until TBLMulta.EOF
            UBMULTAS.Rows = UBMULTAS.Rows + 1
            'UBMULTAS.TextMatrix(f, 0) = TBLMulta!cedula
            'UBMULTAS.TextMatrix(f, 1) = TBLMulta!nombre
            UBMULTAS.TextMatrix(contarfilas, 2) = TBLMulta!fecha
            UBMULTAS.TextMatrix(contarfilas, 3) = TBLMulta!nombre_multa
            UBMULTAS.TextMatrix(contarfilas, 4) = TBLMulta!valor
            UBMULTAS.TextMatrix(contarfilas, 5) = IIf(IsNull(TBLMulta!observacion), "", TBLMulta!observacion)
            UBMULTAS.TextMatrix(contarfilas, 6) = IIf(IsNull(TBLMulta!idsocio_multas), "", TBLMulta!idsocio_multas)
            TBLMulta.MoveNext
            contarfilas = contarfilas + 1
    Loop

End Sub
Private Sub UBMULTAS_Click()
  Dim X As Integer
    VALORPAGAR = 0
    deuda = 0
    Call CONTARSELECCIONADO
    If CONTARREGISTRO < 6 Then
    For X = 1 To UBMULTAS.Rows
    
        If UBMULTAS.TextMatrix(X, 1) = 1 Then
        
        VALORPAGAR = VALORPAGAR + Val(UBMULTAS.TextMatrix(X, 4))
        
        Else
        
        
        deuda = deuda + Val(UBMULTAS.TextMatrix(X, 4))
        
        End If
        
    Next
    Else
        MsgBox "Alcanzado el limete de seleccion", vbInformation, "Cobro Multas"
    End If
    lblValor.Caption = VALORPAGAR
    Me.LBLDEUDA.Caption = deuda
End Sub
Private Sub enviar_a_excel()
    Dim fso, f, i, objeto1
    Dim cadena As String
    Dim NUM As Integer
    Dim Objeto As Object
    
    Set Objeto = Nothing
    Set objeto1 = CreateObject("Excel.Application")
    objeto1.Visible = True
    objeto1.workbooks.Open FileName:=App.Path & "\Recibos\RECIBO1.xlsx"
    cadena = App.Path & "\Recibos\RECIBO1.xlsx"
    
     objeto1.cells(4, 4).Value = Trim(Str(Date))
   ' Set Objeto = GetObject(cadena)
    'Objeto.Application.Windows("RECIBO.xlsx").Visible = True
    
    With objeto1 '.worksheets("RecibosMultas")
    
    
    .cells(4, 4).Value = Trim(Str(Date))
    .cells(8, 2).Value = txtcedula.Text
    .cells(9, 2).Value = Me.txtnombre.Text & " " & Me.txtapellido.Text
    .cells(19, 4).Value = Me.lblValor.Caption
    .cells(2, 5).Value = Me.LblNumeroRecibo.Caption
    
    
    NUM = 13
    Dim cod As Integer
    cod = 1
    For i = 1 To UBMULTAS.Rows
        If UBMULTAS.TextMatrix(i, 1) = 1 Then
        .cells(NUM, 1).Value = cod
        .cells(NUM, 2).Value = UBMULTAS.TextMatrix(i, 3)
        .cells(NUM, 3).Value = UBMULTAS.TextMatrix(i, 5)
        .cells(NUM, 4).Value = UBMULTAS.TextMatrix(i, 4)
        
        cod = cod + 1
        NUM = NUM + 1
        End If
        'insertar fila en excel'
        'Objeto.Range(LTrim(Str(NUM)) & ":" & LTrim(Str(NUM + 1))).Select
        'Objeto.Selection.Insert
        
        
        
    
    Next
    
    
    End With
End Sub
Private Sub CONTARSELECCIONADO()
    'Dim n As Integer
    CONTARREGISTRO = 1
    For i = 1 To UBMULTAS.Rows
        If UBMULTAS.TextMatrix(i, 1) = 1 Then
        CONTARREGISTRO = CONTARREGISTRO + 1
        'NUM = NUM + 1
        End If
        'insertar fila en excel'
        'objeto.Range(LTrim(Str(NUM)) & ":" & LTrim(Str(NUM + 1))).Select
        'objeto.Selection.Insert
        
        
        
    
    Next
End Sub

Private Sub ObtenerNumeroRecibo()
    'para extraer el numero de factura desde la base de datos configuración
    Dim tblConfiguracion As New ADODB.Recordset
        Set tblConfiguracion = Nothing
        tblConfiguracion.Open "SELECT numero FROM numero_recibo", CONEXION, adOpenDynamic, adLockOptimistic
    If Not (tblConfiguracion.EOF) Then
       tblConfiguracion.MoveFirst
       numeroRecibo = tblConfiguracion!numero
        LblNumeroRecibo.Caption = numeroRecibo
   'AQUI indicar el label donde se mostrara el numero---------------
    End If
    Set tblConfiguracion = Nothing
End Sub

Private Sub GenerarNUEVONumeroRecibo()
   'GUARDA EN LA BD EL NUEVO NUMERO DE FACTURA
    Dim tblConfiguracion As New ADODB.Recordset
        Set tblConfiguracion = Nothing
        tblConfiguracion.Open "SELECT numero FROM numero_recibo", CONEXION, adOpenDynamic, adLockOptimistic
    If Not (tblConfiguracion.EOF) Then
       tblConfiguracion.MoveFirst
       tblConfiguracion!numero = numeroRecibo + 1
       tblConfiguracion.Update
       Set tblConfiguracion = Nothing
    End If
    Set tblConfiguracion = Nothing

End Sub

