VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Begin VB.Form Frmregistroaportaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Aportaciones"
   ClientHeight    =   5535
   ClientLeft      =   11535
   ClientTop       =   1935
   ClientWidth     =   8430
   Icon            =   "Frmregistroaportaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
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
      Height          =   975
      Left            =   4320
      Picture         =   "Frmregistroaportaciones.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
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
      Left            =   2280
      Picture         =   "Frmregistroaportaciones.frx":1A90
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
   Begin VB.ComboBox CmbAportaciones 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin ubGridControl.ubGrid UBaportaciones 
      Height          =   1935
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3413
      Rows            =   1
      Cols            =   4
      Redraw          =   -1  'True
      ShowGrid        =   -1  'True
      GridSolid       =   -1  'True
      GridLineColor   =   12632256
      BackColorFixed  =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontEdit {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListBoxRows     =   4
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111869953
      CurrentDate     =   43348
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
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Aportaciones"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Frmregistroaportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLaportaciones As New ADODB.Recordset
Dim contarfilas As Integer
Dim TBLsociomulta1 As New ADODB.Recordset
Private Sub Cmdguardar_Click()
If Len(Me.CmbAportaciones.Text) > 0 Then
'If Len(Cmbmultas1.Text) > 1 And Me.DTPFecha.Value = False Then

    Dim tblguardarsociomulta1 As New ADODB.Recordset
    
     ced = UBaportaciones.TextMatrix(c, 2)
     IDSOCIO = buscarID("socio", "idsocio", IDSOCIO1)
     aportaciones_seleccionada = buscarID("aportaciones", "aportacion", Me.CmbAportaciones.Text)
     Dim z As Integer
     z = MsgBox("Estan marcaddos", vbYesNo, "registros de Aportaciones")
     Dim X As Integer
    For X = 1 To UBaportaciones.Rows
        If UBaportaciones.TextMatrix(X, 1) = 1 Then
             ced = UBaportaciones.TextMatrix(X, 2)
             Set tblguardarsociomulta1 = Nothing
             tblguardarsociomulta1.Open "select * from socio_aportaciones", CONEXION, adOpenDynamic, adLockOptimistic
             tblguardarsociomulta1.AddNew
             tblguardarsociomulta1!IDSOCIO = ced
             tblguardarsociomulta1!idaportaciones = aportaciones_seleccionada
             'tblguardarsociomulta1!observacion = Me.txtObservacion.Text
             tblguardarsociomulta1!fecha = Me.DTPFecha.Value
            'tblguardarsociomulta1!valor = Me.txtValor.Text
            tblguardarsociomulta1!estado_pago = "DEBE"
            
            tblguardarsociomulta1.Update
        
        End If
        
    Next
    'For X = 1 To UBaportaciones.Rows
             'If UBaportaciones.TextMatrix(X, 1) = 0 Then
                Else
            MsgBox "No se puede guardar campos vacios", vbQuestion, "Registro de aportaciones"
        'End If
End If
'Next
       
       'Me.CmbAportaciones.Enabled = False
       UBaportaciones.Clear
       CmbAportaciones.Enabled = False

End Sub
Private Sub CmdRegresar_Click()
    frmAportaciones.Show
    Unload Me
 
End Sub
Private Sub Form_Load()
     contarfilas = 1
    ModuloBaseDatos.conectardb
    Call CargarComboaportaciones
    Call configurargridaportaciones
    Call cargargridaportaciones
    DTPFecha.MaxDate = Date
    DTPFecha.MinDate = Date
End Sub

Private Sub configurargridaportaciones()
    UBaportaciones.AutoSetup 2, 7, True, True, "Marcar|Codigo|cedula|nombre|apellido|"
    
        UBaportaciones.ColMask(1) = checkmark
        'UBMULTAS.ColMask(4) = NumericOnly
        'UBMULTAS.ColMask(6) = NumericOnly
        
    UBaportaciones.ColWidth(1) = 40
    UBaportaciones.ColWidth(2) = 40
    UBaportaciones.ColWidth(3) = 80
    UBaportaciones.ColWidth(4) = 160
    UBaportaciones.ColWidth(5) = 70
    
    
    'UBMULTAS.AutoRedraw = False

    UBaportaciones.ColAllowEdit(2) = False
    UBaportaciones.ColAllowEdit(3) = False
    UBaportaciones.ColAllowEdit(4) = False
    UBaportaciones.ColAllowEdit(5) = False
    
End Sub
Private Sub cargargridaportaciones()
    Dim sql As String
            sql = "select * from socio order by idsocio asc"
           
            Set TBLaportaciones = CONEXION.Execute(sql)
    Dim f As Integer
    If contarfilas > 1 Then
        With UBaportaciones
       ' .Rows = .Rows + 1 ' significa la cantidad de fila
       ' .Row = .Rows - 1 ' significa la fila
        End With
    End If
    
    f = 1
    UBaportaciones.Rows = 0
        Do Until TBLaportaciones.EOF
            UBaportaciones.Rows = UBaportaciones.Rows + 1
            'UBMULTAS.TextMatrix(f, 0) = TBLMulta!cedula
            'UBMULTAS.TextMatrix(f, 1) = TBLMulta!nombre
            UBaportaciones.TextMatrix(contarfilas, 2) = TBLaportaciones!IDSOCIO
            UBaportaciones.TextMatrix(contarfilas, 3) = TBLaportaciones!cedula
           UBaportaciones.TextMatrix(contarfilas, 4) = IIf(IsNull(TBLaportaciones!nombre), "", TBLaportaciones!nombre)
            UBaportaciones.TextMatrix(contarfilas, 5) = IIf(IsNull(TBLaportaciones!apellido), "", TBLaportaciones!apellido)
            TBLaportaciones.MoveNext
            contarfilas = contarfilas + 1
    Loop
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmAportaciones.Enabled = True
    Unload Me
End Sub

Private Sub UBaportaciones_Click()
 For X = 1 To UBaportaciones.Rows
        If UBaportaciones.TextMatrix(X, 1) = 1 Then
        End If
        
    Next
End Sub

Private Sub CargarComboaportaciones()
    CmbAportaciones.Clear
    Set TBLaportaciones = Nothing
    TBLaportaciones.Open "select * from aportaciones", CONEXION, adOpenDynamic, adLockOptimistic
    Do Until TBLaportaciones.EOF
    CmbAportaciones.AddItem TBLaportaciones.Fields(2).Value
        TBLaportaciones.MoveNext
        Loop
End Sub
