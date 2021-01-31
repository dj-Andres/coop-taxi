VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Begin VB.Form FrmCobranzasAportaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobranzas Aportaciones"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmCobranzasAporatciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   720
      TabIndex        =   17
      Top             =   7560
      Width           =   8415
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   1300
         Left            =   3360
         Picture         =   "FrmCobranzasAporatciones.frx":10CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1500
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   1300
         Left            =   1800
         Picture         =   "FrmCobranzasAporatciones.frx":554A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   1300
         Left            =   5040
         Picture         =   "FrmCobranzasAporatciones.frx":9400
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1500
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   1300
         Left            =   6720
         Picture         =   "FrmCobranzasAporatciones.frx":D0CB
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1500
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Cobrar"
         Height          =   1300
         Left            =   120
         Picture         =   "FrmCobranzasAporatciones.frx":1119B
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.TextBox txtObservacion 
      Height          =   690
      Left            =   2520
      MaxLength       =   25
      TabIndex        =   15
      Top             =   6000
      Width           =   5895
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   112394241
      CurrentDate     =   43380
   End
   Begin VB.CommandButton cmdBuscar 
      Height          =   615
      Left            =   5280
      Picture         =   "FrmCobranzasAporatciones.frx":13D75
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdCargarAportaciones 
      Caption         =   "Cargar Aportaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox txtapellido 
      Height          =   330
      Left            =   2520
      TabIndex        =   7
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtnombre 
      Height          =   330
      Left            =   2520
      TabIndex        =   6
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtcedula 
      Height          =   330
      Left            =   2520
      TabIndex        =   5
      Top             =   1320
      Width           =   2535
   End
   Begin ubGridControl.ubGrid UBAPORTACIONES 
      Height          =   1575
      Left            =   720
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
   Begin VB.Label Label9 
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
      Left            =   6600
      TabIndex        =   22
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label LblNumeroRecibo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   7800
      TabIndex        =   23
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label lblValor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   515
      Left            =   2520
      TabIndex        =   16
      Top             =   6840
      Width           =   600
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
      Left            =   720
      TabIndex        =   13
      Top             =   6840
      Width           =   510
   End
   Begin VB.Label Label7 
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
      Left            =   720
      TabIndex        =   12
      Top             =   6120
      Width           =   1155
   End
   Begin VB.Label Label6 
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
      Left            =   720
      TabIndex        =   11
      Top             =   5520
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aportaciones"
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
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Left            =   720
      TabIndex        =   3
      Top             =   2280
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
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   765
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
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COBRANZAS DE APORTACIONES"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   7170
   End
End
Attribute VB_Name = "FrmCobranzasAportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLAportacion As New ADODB.Recordset
Dim TBLGuardarCobro As New ADODB.Recordset
Dim VALORPAGAR As Double
Dim contarfilas As Integer
Dim TBLCobranzas As New ADODB.Recordset
Dim deuda As Double
Dim idcobranza_aportacion1 As Integer
Dim tblsocioaportacion As New ADODB.Recordset
Dim numeroRecibo As Integer
Dim ok As Boolean
Dim BAN_imprimir As String

Private Sub cmdAgregar_Click()
    Call activarcajas
    frmBusquedaCobranzasAportaciones.Show
    FrmCobranzasAportaciones.Enabled = False
    Call ObtenerNumeroRecibo
           
End Sub

Private Sub cmdBuscar_Click()
    frmBusquedaCobranzasAportaciones.Show
End Sub
Private Sub cmdCancelar_Click()
    Call desactivarcajas
End Sub
Private Sub cmdCargarAportaciones_Click()
    Call configurargridAPORTACION
    Call cargargridAPORTACION
    contarfilas = 1
    Me.UBAPORTACIONES.Enabled = True
End Sub

Private Sub cmdCerrar_Click()
frmMenu.Show
FrmCobranzasAportaciones.Hide
End Sub

Private Sub Cmdguardar_Click()
    'Dim tblguardarsociomulta1 As New ADODB.Recordset
    'Dim socio_ingresado As Integer
    'Dim multa_seleccionada As Integer
     'Dim observacion_ingresada As Integer
     'Dim valor_ingresado As Integer
     'observacion_ingresada = buscarID("multas", "observacion", Me.txtObservacion.Text)
     'valor_ingresado = buscarID("multas", "valor", Me.txtValor.Text)
     
     socio_aportacion = buscarID("socio", "idsocio", IDSOCIO1)
     'multa_seleccionada = buscarID("multas", "nombre_multa", Me.cmbMultas.Text)
        
        Set TBLGuardarCobro = Nothing
        TBLGuardarCobro.Open "select * from cobranzas_aportaciones", CONEXION, adOpenDynamic, adLockOptimistic
        TBLGuardarCobro.AddNew
        TBLGuardarCobro!IDSOCIO = socio_aportacion
        TBLGuardarCobro!fecha = DTPFecha.Value
        TBLGuardarCobro!observacion = Me.txtObservacion.Text
        TBLGuardarCobro!total_pagar = Me.lblValor.Caption
        TBLGuardarCobro!idusuario = 1 ' IDUSUARIO1
        'TBLGuardarCobro!saldo = Me.LBLDEUDA.Caption
        TBLGuardarCobro!numero_recibo = Me.LblNumeroRecibo.Caption

        
         
        TBLGuardarCobro.Update
        TBLGuardarCobro.Requery
        
        Set TBLGuardarCobro = Nothing
        TBLGuardarCobro.Open "select * from cobranzas_aportaciones", CONEXION, adOpenDynamic, adLockOptimistic
        TBLGuardarCobro.MoveLast
        idcobranza_aportacion1 = TBLGuardarCobro!idcobranzas_aportaciones
        Set TBLGuardarCobro = Nothing
        
        
        
        Set tblsociaportacion = Nothing
        
        
        Dim sql As String
           
        Dim a As Integer
        
        For a = 1 To UBAPORTACIONES.Rows
        If UBAPORTACIONES.TextMatrix(a, 1) = 1 Then
            sql = "update  socio_aportaciones set estado_pago='CANCELADO' where idsocio_aportaciones='" & UBAPORTACIONES.TextMatrix(a, 6) & "'"
            Set TBLAportacion = CONEXION.Execute(sql)
               sql = "update  socio_aportaciones set referencia='" & idcobranza_aportacion1 & "' where idsocio_aportaciones='" & UBAPORTACIONES.TextMatrix(a, 6) & "'"
            Set tblsocioaportacion = CONEXION.Execute(sql)
               
        End If
    Next
        MsgBox "EL PAGO HA SIDO GUARDADO EXITOSAMENTE", vbInformation, "COBROS DE MULTAS"
    
    cmdImprimir.Enabled = True
    cmdGuardar.Enabled = False
    cmdGuardar.Enabled = False
    UBAPORTACIONES.Enabled = False
    DTPFecha.Enabled = False
    txtObservacion.Enabled = False
    
    Call GenerarNUEVONumeroRecibo


End Sub
Private Sub cmdImprimir_Click()
    Call enviar_a_excel
    UBAPORTACIONES.Clear
    txtcedula.Text = ""
    txtnombre.Text = ""
    txtapellido.Text = ""
    DTPFecha.Enabled = False
    Me.txtObservacion.Text = ""
    Me.lblValor.Caption = ""

End Sub
Private Sub Form_Load()
    ModuloBaseDatos.conectardb
    Call desactivarcajas
    Me.DTPFecha.MaxDate = Date
    Me.DTPFecha.MinDate = Date
End Sub
Private Sub configurargridAPORTACION()
    UBAPORTACIONES.AutoSetup 2, 7, True, True, "Marcar|fecha|Aportacion|valor|Descripcion|Codigo"
    
        UBAPORTACIONES.ColMask(1) = checkmark
        UBAPORTACIONES.ColMask(4) = NumericOnly
        UBAPORTACIONES.ColMask(6) = NumericOnly
        
    UBAPORTACIONES.ColWidth(1) = 40
    UBAPORTACIONES.ColWidth(2) = 80
    UBAPORTACIONES.ColWidth(3) = 160
    UBAPORTACIONES.ColWidth(4) = 100
    UBAPORTACIONES.ColWidth(5) = 150
    UBAPORTACIONES.ColWidth(6) = 30
    
    'UBMULTAS.AutoRedraw = False

    UBAPORTACIONES.ColAllowEdit(2) = False
    UBAPORTACIONES.ColAllowEdit(3) = False
    UBAPORTACIONES.ColAllowEdit(4) = False
    UBAPORTACIONES.ColAllowEdit(5) = False
    UBAPORTACIONES.ColAllowEdit(6) = False
     
End Sub
Private Sub cargargridAPORTACION()
    Dim sql As String
            sql = "select sa.fecha,a.aportacion,a.valor,a.descripcion,sa.idsocio_aportaciones,sa.estado_pago from socio_aportaciones sa join aportaciones a on a.idaportaciones=sa.idaportaciones where sa.estado_pago= 'DEBE'  and idsocio='" & IDSOCIO1 & "'"
           
            Set TBLAportacion = CONEXION.Execute(sql)
    Dim f As Integer
    If contarfilas > 1 Then
        With UBAPORTACIONES
       ' .Rows = .Rows + 1 ' significa la cantidad de fila
       ' .Row = .Rows - 1 ' significa la fila
        End With
    End If
    
    f = 1
    UBAPORTACIONES.Rows = 0
        Do Until TBLAportacion.EOF
            UBAPORTACIONES.Rows = UBAPORTACIONES.Rows + 1
            'UBMULTAS.TextMatrix(f, 0) = TBLMulta!cedula
            'UBMULTAS.TextMatrix(f, 1) = TBLMulta!nombre
            UBAPORTACIONES.TextMatrix(contarfilas, 2) = TBLAportacion!fecha
            UBAPORTACIONES.TextMatrix(contarfilas, 3) = IIf(IsNull(TBLAportacion!aportacion), "", TBLAportacion!aportacion)
            UBAPORTACIONES.TextMatrix(contarfilas, 4) = TBLAportacion!valor
            UBAPORTACIONES.TextMatrix(contarfilas, 5) = IIf(IsNull(TBLAportacion!descripcion), "", TBLAportacion!descripcion)
            UBAPORTACIONES.TextMatrix(contarfilas, 6) = IIf(IsNull(TBLAportacion!idsocio_aportaciones), "", TBLAportacion!idsocio_aportaciones)
            TBLAportacion.MoveNext
            contarfilas = contarfilas + 1
    Loop

End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMenu.Show
    Unload Me
End Sub


Private Sub UBaportaciones_Click()
      Dim X As Integer
    VALORPAGAR = 0
    'deuda = 0
    
    For X = 1 To UBAPORTACIONES.Rows
        If UBAPORTACIONES.TextMatrix(X, 1) = 1 Then
        
        VALORPAGAR = VALORPAGAR + Val(UBAPORTACIONES.TextMatrix(X, 4))
        
        'Else
        
        
        'deuda = deuda + Val(UBAPORTACIONES.TextMatrix(x, 4))
        
        End If
        
    Next
    lblValor.Caption = VALORPAGAR
    'Me.LBLDEUDA.Caption = deuda

End Sub
Private Sub enviar_a_excel()
    Dim fso, f, i, objeto1
    Dim cadena As String
    Dim NUM As Integer
    Dim Objeto As Object
    
    Set Objeto = Nothing
    Set objeto1 = CreateObject("Excel.Application")
    objeto1.Visible = True
    objeto1.workbooks.Open FileName:=App.Path & "\Recibos\RECIBO2.xlsx"
    cadena = App.Path & "\Recibos\RECIBO2.xlsx"
  
    ok = BUSCA_ARCHIVO(cadena)
      
  If Not ok Then
         MsgBox "ERROR: El Archivo de Impresión no existe ", vbOKOnly + vbCritical, "ARCHIVO"
        Else
         Set fso = CreateObject("Scripting.FileSystemObject")
         Set COPIA = CreateObject("Scripting.FileSystemObject")
         If (fso.FolderExists("C:\ARCHIVOS RECIBOS")) Then
            COPIA.COPYFile cadena, "C:\ARCHIVOS RECIBOS\"
           Else
            Set f = fso.CreateFolder("C:\ARCHIVOS RECIBOS")
            COPIA.COPYFile cadena, "C:\ARCHIVOS RECIBOS\"
         End If
    
    
Set Objeto = CreateObject("excel.Application")
    
     objeto1.cells(4, 4).Value = Trim(Str(Date))
   ' Set Objeto = GetObject(cadena)
    'Objeto.Application.Windows("RECIBO.xlsx").Visible = True
    
    With objeto1 '.worksheets("RecibosMultas")
    
    
    .cells(4, 4).Value = Trim(Str(Date))
    .cells(8, 2).Value = txtcedula.Text
    .cells(9, 2).Value = Me.txtnombre.Text & " " & Me.txtapellido.Text
    .cells(17, 4).Value = Me.lblValor.Caption
    .cells(2, 5).Value = Me.LblNumeroRecibo.Caption
    
    NUM = 12
    Dim cod As Integer
    cod = 1
    For i = 1 To UBAPORTACIONES.Rows
        If UBAPORTACIONES.TextMatrix(i, 1) = 1 Then
        .cells(NUM, 1).Value = i
        .cells(NUM, 2).Value = UBAPORTACIONES.TextMatrix(i, 3)
        .cells(NUM, 3).Value = UBAPORTACIONES.TextMatrix(i, 5)
        .cells(NUM, 4).Value = UBAPORTACIONES.TextMatrix(i, 4)
        
        'insertar fila en excel'
        'objeto.Range(LTrim(Str(NUM)) & ":" & LTrim(Str(NUM + 1))).Select
        'objeto.Selection.Insert
        cod = cod + 1
        NUM = NUM + 1
        
        End If
    
    Next
    
    
    End With
    
        Set Objeto = Nothing
End If

End Sub
Private Sub desactivarcajas()
    txtcedula.Enabled = False
    txtnombre.Enabled = False
    txtapellido.Enabled = False
    txtObservacion.Enabled = False
    Me.lblValor.Enabled = False
    DTPFecha.Enabled = False
    Me.cmdCargarAportaciones.Enabled = False
    cmdBuscar.Enabled = False
    
    Me.UBAPORTACIONES.Enabled = False
    
    txtObservacion.Text = ""
    txtcedula.Text = ""
    txtnombre.Text = ""
    txtapellido.Text = ""
    
    Me.cmdCerrar.Enabled = True
    cmdAgregar.Enabled = True
    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False
    cmdImprimir.Enabled = False
End Sub
Private Sub activarcajas()
    txtcedula.Enabled = False
    txtnombre.Enabled = False
    txtapellido.Enabled = False
    txtObservacion.Enabled = True
    Me.lblValor.Enabled = True
    DTPFecha.Enabled = True
    Me.cmdCargarAportaciones.Enabled = True
    cmdBuscar.Enabled = True
    
    txtObservacion.Text = ""
    
    Me.cmdCerrar.Enabled = True
    cmdAgregar.Enabled = False
    cmdGuardar.Enabled = True
    cmdCancelar.Enabled = True
    cmdImprimir.Enabled = True

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
   'AQUI indicar el label donde se mostrara el nuermo---------------
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

Public Function BUSCA_ARCHIVO(ByVal cadena As String) As Boolean
    Dim ARCHIVO2
    Set ARCHIVO2 = CreateObject("Scripting.FileSystemObject")
    BUSCA_ARCHIVO = ARCHIVO2.FileExists(cadena)
End Function



