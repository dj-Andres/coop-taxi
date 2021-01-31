VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Begin VB.Form Registro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Multas"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9810
   Icon            =   "Frmregistro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Frmregistro.frx":424A
   ScaleHeight     =   4860
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "Regresar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      Picture         =   "Frmregistro.frx":458C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Cmdguardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      Picture         =   "Frmregistro.frx":4F52
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   2055
   End
   Begin VB.ComboBox Cmbmultas1 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Text            =   "cmbmultas1"
      Top             =   1080
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111673345
      CurrentDate     =   43348
   End
   Begin ubGridControl.ubGrid UBMULTAS 
      Height          =   1575
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Multas"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   600
   End
End
Attribute VB_Name = "Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLMulta As New ADODB.Recordset
Dim contarfilas As Integer
Dim TBLsociomulta1 As New ADODB.Recordset
Private Sub CargarCombomultas1()
    Cmbmultas1.Clear
    Set TBLMulta = Nothing
    TBLMulta.Open "select * from multas", CONEXION, adOpenDynamic, adLockOptimistic
    Do Until TBLMulta.EOF
    Cmbmultas1.AddItem TBLMulta.Fields(1).Value
        TBLMulta.MoveNext
        Loop
End Sub
Private Sub Cmdguardar_Click()
If Len(Cmbmultas1.Text) > 0 Then
Dim IDSOCIO As Integer
    Dim tblguardarsociomulta1 As New ADODB.Recordset
    'Dim socio_ingresado As Integer
    'Dim multa_seleccionada As Integer
     'Dim observacion_ingresada As Integer
     'Dim valor_ingresado As Integer
     'observacion_ingresada = buscarID("multas", "observacion", Me.txtObservacion.Text)
     'valor_ingresado = buscarID("multas", "valor", Me.txtValor.Text)
     
     'socio_ingresado = buscarID("socio", "idsocio", IDSOCIO1)
     ced = UBMULTAS.TextMatrix(c, 2)
     IDSOCIO = buscarID("socio", "idsocio", IDSOCIO1)
     multa_seleccionada = buscarID("multas", "nombre_multa", Me.Cmbmultas1.Text)
     Dim z As Integer
     z = MsgBox("Estan marcados", vbYesNo, "registros de Multas")
     Dim X As Integer
    For X = 1 To UBMULTAS.Rows
        If UBMULTAS.TextMatrix(X, 1) = 1 Then
             ced = UBMULTAS.TextMatrix(X, 2)
             Set tblguardarsociomulta1 = Nothing
             tblguardarsociomulta1.Open "select * from socio_multas", CONEXION, adOpenDynamic, adLockOptimistic
             tblguardarsociomulta1.AddNew
             tblguardarsociomulta1!IDSOCIO = ced
             tblguardarsociomulta1!idmulta = multa_seleccionada
             'tblguardarsociomulta1!observacion = Me.txtObservacion.Text
             tblguardarsociomulta1!fecha = Me.DTPFecha.Value
            'tblguardarsociomulta1!valor = Me.txtValor.Text
            tblguardarsociomulta1!estado_pago = "DEBE"
            
            tblguardarsociomulta1.Update
            
         
        End If
        
   Next
    'For X = 1 To UBMULTAS.Rows
         'UBMULTAS.TextMatrix(X, 1) = 0
         
    Else
        MsgBox "No se puede guardar campos vacios", vbQuestion, "Registro de Multas"
       'Else
      ' Me.Cmbmultas1.Enabled = True
       'Call cargargridMULTAS
       

      
        
        End If
        UBMULTAS.Clear
       Cmbmultas1.Text = ""
       
    
        
                
End Sub

Private Sub CmdRegresar_Click()
    frmGestionMultas.Show
    Unload Me
    
    
End Sub

Private Sub Form_Load()
    contarfilas = 1
    ModuloBaseDatos.conectardb
    Call CargarCombomultas1
    Call configurargridMULTAS
    Call cargargridMULTAS
    DTPFecha.MaxDate = Date
    DTPFecha.MinDate = Date

End Sub
Private Sub configurargridMULTAS()
    UBMULTAS.AutoSetup 2, 7, True, True, "Marcar|Codigo|cedula|nombre|apellido|"
    
        UBMULTAS.ColMask(1) = checkmark
        'UBMULTAS.ColMask(4) = NumericOnly
        'UBMULTAS.ColMask(6) = NumericOnly
        
   UBMULTAS.ColWidth(1) = 70
    UBMULTAS.ColWidth(2) = 60
    UBMULTAS.ColWidth(3) = 90
    UBMULTAS.ColWidth(4) = 160
    UBMULTAS.ColWidth(5) = 160
    
    
    'UBMULTAS.AutoRedraw = False

    UBMULTAS.ColAllowEdit(2) = False
    UBMULTAS.ColAllowEdit(3) = False
    UBMULTAS.ColAllowEdit(4) = False
    UBMULTAS.ColAllowEdit(5) = False
    
End Sub
Private Sub cargargridMULTAS()
    Dim sql As String
            sql = "select * from socio"
           
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
            UBMULTAS.TextMatrix(contarfilas, 2) = IIf(IsNull(TBLMulta!IDSOCIO), "", TBLMulta!IDSOCIO)
            UBMULTAS.TextMatrix(contarfilas, 3) = TBLMulta!cedula
            UBMULTAS.TextMatrix(contarfilas, 4) = IIf(IsNull(TBLMulta!nombre), "", TBLMulta!nombre)
            UBMULTAS.TextMatrix(contarfilas, 5) = IIf(IsNull(TBLMulta!apellido), "", TBLMulta!apellido)
            TBLMulta.MoveNext
            
            contarfilas = contarfilas + 1
    Loop

End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmGestionMultas.Enabled = True
    Unload Me
End Sub

Private Sub UBMULTAS_Click()
    For X = 1 To UBMULTAS.Rows
        If UBMULTAS.TextMatrix(X, 1) = 1 Then
        End If
        
    Next
End Sub
