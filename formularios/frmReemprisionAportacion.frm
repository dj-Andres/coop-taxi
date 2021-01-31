VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReemprisionAportacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reimprimir Aportacion"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9255
   Icon            =   "frmReemprisionAportacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1335
      Left            =   4080
      Picture         =   "frmReemprisionAportacion.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Reimprimir"
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
      Left            =   2160
      Picture         =   "frmReemprisionAportacion.frx":1290
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1500
   End
   Begin VB.TextBox txtbusqueda 
      Height          =   495
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
   Begin MSFlexGridLib.MSFlexGrid GridDetalle 
      Height          =   1335
      Left            =   600
      TabIndex        =   2
      Top             =   3000
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2355
      _Version        =   393216
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid GridRecibo 
      Height          =   1335
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2355
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
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
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Width           =   1245
   End
End
Attribute VB_Name = "frmReemprisionAportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim contarfilas As Integer
Dim cedula As Integer
Dim cisocio As Integer
Dim TBLSocio As New ADODB.Recordset

Private Sub cmdCerrar_Click()
frmMenu.Show
Unload Me
End Sub

Private Sub cmdImprimir_Click()
     Call enviar_a_excel
End Sub

Private Sub Form_Load()
    ModuloBaseDatos.conectardb
    Call configurargrid
    Call cargargrid
    Call configurarGRIDdetalle
    contarfilas = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMenu.Show
    Unload Me
End Sub
Private Sub cargargrid()
     
     Dim sql As String
    sql = "   select s.cedula,s.nombre,s.apellido,ca.fecha,ca.numero_recibo from socio s join socio_aportaciones sa  on s.idsocio=sa.idsocio join aportaciones a on  a.idaportaciones = sa.idaportaciones join cobranzas_aportaciones ca  on ca.idcobranzas_aportaciones=sa.referencia where estado_pago='CANCELADO'"
        Set TBLSocio = CONEXION.Execute(sql)
    Dim f As Integer
    f = 1
    GridRecibo.Rows = 2
    Do Until TBLSocio.EOF
        'GridRecibo.TextMatrix(f, 0) = TblSocio!IDSOCIO
        GridRecibo.TextMatrix(f, 0) = TBLSocio!cedula
        GridRecibo.TextMatrix(f, 1) = TBLSocio!nombre
        GridRecibo.TextMatrix(f, 2) = TBLSocio!apellido
        GridRecibo.TextMatrix(f, 3) = TBLSocio!fecha
        'GridRecibo.TextMatrix(f, 4) = TBLSocio!total_pagar
        GridRecibo.TextMatrix(f, 4) = IIf(IsNull(TBLSocio!numero_recibo), "", TBLSocio!numero_recibo)
        'GridRecibo.TextMatrix(f, 6) = TblSocio!valor
        
        TBLSocio.MoveNext
        f = f + 1
        GridRecibo.Rows = GridRecibo.Rows + 1
        
    Loop

End Sub
Private Sub configurargrid()
                                                                                                                                                     
    GridRecibo.Clear
    GridRecibo.FormatString = "cedula|nombre|apellido|fecha|numero recibo"
    GridRecibo.ColWidth(0) = 1000
    GridRecibo.ColWidth(1) = 1500
    GridRecibo.ColWidth(2) = 1500
    GridRecibo.ColWidth(3) = 1200
    GridRecibo.ColWidth(4) = 700
    'GridRecibo.ColWidth(5) = 1000
    'GridRecibo.ColWidth(6) = 1500
    'GridRecibo.ColWidth(7) = 1500
    

End Sub
Private Sub GridRecibo_DblClick()
Dim sql As String
'Dim f As Integer
        cisocio = GridRecibo.TextMatrix(GridRecibo.Row, 4)
        cedula = ModuloFunciones.buscarID("cobranzas_aportaciones", "numero_recibo", cisocio)
        
        Set TBLSocio = Nothing
             TBLSocio.Open "select a.aportacion,ca.fecha,ca.observacion,ca.total_pagar from socio s join socio_aportaciones sa on s.idsocio=sa.idsocio join aportaciones a on  a.idaportaciones = sa.idaportaciones join cobranzas_aportaciones ca  on ca.idcobranzas_aportaciones=sa.referencia  where estado_pago='CANCELADO' and numero_recibo='" & cisocio & "'", CONEXION, adOpenDynamic, adLockOptimistic
            
            TBLSocio.MoveFirst
            
            If contarfilas > 1 Then
             With GridDetalle
                .Rows = .Rows + 1
                .Row = .Rows - 1
             
             End With
            End If
             
            'GridDetalle.TextMatrix(contarfilas, 0) = GridRecibo.TextMatrix(contarfilas, 0)
            'GridDetalle.TextMatrix(contarfilas, 1) = GridRecibo.TextMatrix(contarfilas, 1)
            'GridDetalle.TextMatrix(contarfilas, 2) = GridRecibo.TextMatrix(contarfilas, 2)
            'GridDetalle.TextMatrix(contarfilas, 3) = GridRecibo.TextMatrix(contarfilas, 3)
            'Call cargarGRIDdetalle1
            'For X = contarfilas To GridDetalle.Rows
                
                GridDetalle.TextMatrix(contarfilas, 0) = TBLSocio!aportacion
                GridDetalle.TextMatrix(contarfilas, 1) = TBLSocio!fecha
                GridDetalle.TextMatrix(contarfilas, 2) = TBLSocio!observacion
                GridDetalle.TextMatrix(contarfilas, 3) = TBLSocio!total_pagar
                TBLSocio.MoveNext
             
        'Next
        
End Sub


Private Sub txtbusqueda_Change()
    Dim sql As String
    sql = "   select  s.cedula,s.nombre,s.apellido,ca.fecha,ca.total_pagar,ca.numero_recibo from socio s join socio_aportaciones sa on s.idsocio=sa.idsocio join aportaciones a on  a.idaportaciones = sa.idaportaciones join cobranzas_aportaciones ca  on ca.idcobranzas_aportaciones=sa.referencia where cedula like '%" & txtbusqueda.Text & "%'  or  nombre like '%" & Trim(UCase(txtbusqueda.Text)) & "%'"
        Set TBLSocio = CONEXION.Execute(sql)
    Dim f As Integer
    f = 1
    GridRecibo.Rows = 2
    Do Until TBLSocio.EOF
        'GridRecibo.TextMatrix(f, 0) = TblSocio!IDSOCIO
        GridRecibo.TextMatrix(f, 0) = TBLSocio!cedula
        GridRecibo.TextMatrix(f, 1) = TBLSocio!nombre
        GridRecibo.TextMatrix(f, 2) = TBLSocio!apellido
        GridRecibo.TextMatrix(f, 3) = TBLSocio!fecha
        'GridRecibo.TextMatrix(f, 4) = TBLSocio!total_pagar
        GridRecibo.TextMatrix(f, 4) = IIf(IsNull(TBLSocio!numero_recibo), "", TBLSocio!numero_recibo)
        'GridRecibo.TextMatrix(f, 6) = TblSocio!valor
        
        TBLSocio.MoveNext
        f = f + 1
        GridRecibo.Rows = GridRecibo.Rows + 1
        
    Loop

End Sub
Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    KeyAscii = ModuloFunciones.numeros_letras(KeyAscii)
End Sub
Private Sub configurarGRIDdetalle()
GridDetalle.Clear
    GridDetalle.FormatString = "Aportacion|fecha|observacion|valor"
    GridDetalle.ColWidth(0) = 2000
    GridDetalle.ColWidth(1) = 1500
    GridDetalle.ColWidth(2) = 1500
    GridDetalle.ColWidth(3) = 1500

End Sub
Private Sub enviar_a_excel()
 Dim fso, f, i, objeto1
    Dim cadena As String
    Dim NUM As Integer
    Dim Objeto As Object
    
    Set Objeto = Nothing
    Set objeto1 = CreateObject("Excel.Application")
    objeto1.Visible = True
    objeto1.workbooks.Open FileName:=App.Path & "\Recibos\ReempresionAPORTACION.xls"
    cadena = App.Path & "\Recibos\ReempresionAPORTACION.xls"
    
     'objeto1.cells(4, 4).Value = Trim(Str(Date))
   ' Set Objeto = GetObject(cadena)
    'Objeto.Application.Windows("RECIBO.xlsx").Visible = True
    
    With objeto1 '.worksheets("RecibosMultas")
    
    
   ' .cells(2, 4).Value = Trim(Str(Date))
    '.cells(1, 4).Value = Me.txtnombre.Text
    '.cells(3, 4).Value = Me.lblValor.Caption
    '.cells(4, 4).Value = MeCaption

    
    NUM = 9
    Dim cod As Integer
    cod = 1
    For i = 1 To GridRecibo.Rows - 1
        'If GridRecibo.TextMatrix(i, 0) = 1 Then
        '.cells(NUM, 1).Value = i
        .cells(9, 1).Value = GridRecibo.TextMatrix(i, 0)
        .cells(9, 2).Value = GridRecibo.TextMatrix(i, 1)
        .cells(9, 3).Value = GridRecibo.TextMatrix(i, 2)
        .cells(9, 4).Value = GridRecibo.TextMatrix(i, 4)
        .cells(12, 1).Value = GridDetalle.TextMatrix(i, 0)
        .cells(12, 2).Value = GridDetalle.TextMatrix(i, 1)
        .cells(12, 3).Value = GridDetalle.TextMatrix(i, 2)
        .cells(12, 4).Value = GridDetalle.TextMatrix(i, 3)
        '.cells(NUM, 5).Value = GridRecibo.TextMatrix(i, 4)
        '.cells(NUM, 6).Value = GridRecibo.TextMatrix(i, 5)

        'insertar fila en excel'
        'objeto.Range(LTrim(Str(NUM)) & ":" & LTrim(Str(NUM + 1))).Select
        'objeto.Selection.Insert
        cod = cod + 1
        NUM = NUM + 1
        
        'End If
    
    Next
    'NUM = 9
     'For i = 1 To GridDetalle.Rows - 1
        'If GridRecibo.TextMatrix(i, 1) = 1 Then
        '.cells(NUM, 1).Value = i
        '.cells(NUM, 5).Value = GridDetalle.TextMatrix(i, 0)
        '.cells(NUM, 6).Value = GridDetalle.TextMatrix(i, 1)
        '.cells(NUM, 7).Value = GridDetalle.TextMatrix(i, 2)
        '.cells(NUM, 8).Value = GridDetalle.TextMatrix(i, 3)
        'insertar fila en excel'
        'objeto.Range(LTrim(Str(NUM)) & ":" & LTrim(Str(NUM + 1))).Select
        'objeto.Selection.Insert
        'cod = cod + 1
        'NUM = NUM + 1
        
        'End If
        'Next
    End With
End Sub


