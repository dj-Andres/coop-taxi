VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAgregarAportaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Aportaciones"
   ClientHeight    =   5820
   ClientLeft      =   12435
   ClientTop       =   5460
   ClientWidth     =   9690
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmAgregarAportaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9690
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   360
      Picture         =   "FrmAgregarAportaciones.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      Width           =   1500
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   360
      Picture         =   "FrmAgregarAportaciones.frx":4FC0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   720
      Width           =   1500
   End
   Begin VB.CommandButton CmdRegresar 
      Caption         =   "Regresar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   360
      Picture         =   "FrmAgregarAportaciones.frx":8E76
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      Begin MSFlexGridLib.MSFlexGrid GridAportacion 
         Height          =   2175
         Left            =   480
         TabIndex        =   10
         Top             =   2280
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin VB.TextBox TxtAportaciones 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   3
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox TxtDescripcion 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   2
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox TxtValor 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label LblAportaciones 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aportaciones"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   2010
      End
      Begin VB.Label LblDsc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   5
         Top             =   960
         Width           =   1560
      End
      Begin VB.Label LblValor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   750
      End
   End
End
Attribute VB_Name = "FrmAgregarAportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLAportacion As New ADODB.Recordset
Dim idaportacion As Integer
Dim ciaportacion As Integer
Dim TBLGuardarAportacion As New ADODB.Recordset
Private Sub Cmdguardar_Click()
    If Len(TxtAportaciones.Text) > 1 And Len(TxtDescripcion.Text) > 1 And Len(TxtValor.Text) > 1 Then
        Dim respuestaID As Integer
        Set TBLGuardarAportacion = Nothing
            TBLGuardarAportacion.Open "select * from aportaciones where aportacion= '" & UCase(TxtAportaciones.Text) & "'", CONEXION, adOpenDynamic, adLockOptimistic
            If Not (TBLGuardarAportacion.EOF) Then
                MsgBox "La aportacion ya existe"
            
            
            Else
                'Dim sql As String
                'sql = "insert into aportaciones (aportacion, descripcion, valor)values('" & UCase(TxtAportaciones.Text) & "', '" & UCase(TxtDescripcion.Text) & "', " & UCase(TxtValor.Text) & ")"
                'MsgBox sql
                Set TBLGuardarAportacion = Nothing
                TBLGuardarAportacion.Open "select * from aportaciones", CONEXION, adOpenDynamic, adLockOptimistic
                TBLGuardarAportacion.AddNew
                TBLGuardarAportacion!valor = Me.TxtValor.Text
                TBLGuardarAportacion!aportacion = Trim(UCase(Me.TxtAportaciones.Text))
                TBLGuardarAportacion!descripcion = Trim(UCase(TxtDescripcion.Text))
                
                TBLGuardarAportacion.Update
                'Call cargargrid
                
                'Set TBLGuardarAportacion = CONEXION.Execute(sql)
                Call CargarAportacion
            End If
    Else
        MsgBox "Ingrese todos los campos", vbInformation, "Agregar Aportacion"
    End If
    TxtAportaciones.Text = ""
    TxtDescripcion.Text = ""
    TxtValor.Text = ""
    
End Sub
Private Sub cmdModificar_Click()
    If Len(TxtAportaciones.Text) > 1 And Len(TxtDescripcion.Text) > 1 And Len(TxtValor.Text) > 1 Then
        Dim respuestaID As Integer
        Set TBLGuardarAportacion = Nothing
            'TBLGuardarAportacion.Open "select * from aportaciones where aportacion= '" & UCase(TxtAportaciones.Text) & "'", CONEXION, adOpenDynamic, adLockOptimistic
            'If Not (TBLGuardarAportacion.EOF) Then
                'MsgBox "La aportacion ya existe"
            
            
            'Else
                'Dim sql As String
                'sql = "insert into aportaciones (aportacion, descripcion, valor)values('" & UCase(TxtAportaciones.Text) & "', '" & UCase(TxtDescripcion.Text) & "', " & UCase(TxtValor.Text) & ")"
                'MsgBox sql
                Set TBLGuardarAportacion = Nothing
                TBLGuardarAportacion.Open "select * from aportaciones where idaportaciones=" & ciaportacion, CONEXION, adOpenDynamic, adLockOptimistic
                'TBLGuardarAportacion.AddNew
                TBLGuardarAportacion!valor = Me.TxtValor.Text
                TBLGuardarAportacion!aportacion = Trim(UCase(Me.TxtAportaciones.Text))
                TBLGuardarAportacion!descripcion = Trim(UCase(TxtDescripcion.Text))
                
                TBLGuardarAportacion.Update
                'Call cargargrid
                
                'Set TBLGuardarAportacion = CONEXION.Execute(sql)
                Call CargarAportacion
            End If
    'Else
        'MsgBox "Ingrese registro"
    'End If
    TxtAportaciones.Text = ""
    TxtDescripcion.Text = ""
    TxtValor.Text = ""
    
End Sub
Private Sub CmdRegresar_Click()
    Unload Me
    frmAportaciones.Enabled = True
End Sub
Private Sub Form_Load()
    Call configurargrid
    Call CargarAportacion
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmAportaciones.Enabled = True
End Sub
Private Sub CargarAportacion()
    Dim sql As String
    sql = "select * from aportaciones"
    Set TBLAportacion = CONEXION.Execute(sql)
    Dim f As Integer
    f = 1
    GridAportacion.Rows = 2
    Do Until TBLAportacion.EOF
        GridAportacion.TextMatrix(f, 0) = TBLAportacion!idaportaciones
        GridAportacion.TextMatrix(f, 1) = TBLAportacion!valor
        GridAportacion.TextMatrix(f, 2) = TBLAportacion!aportacion
        GridAportacion.TextMatrix(f, 3) = IIf(IsNull(TBLAportacion!descripcion), "", TBLAportacion!descripcion)
        TBLAportacion.MoveNext
        f = f + 1
        GridAportacion.Rows = GridAportacion.Rows + 1
    Loop
End Sub
Private Sub GridAportacion_DblClick()
    Dim z As Integer
    z = MsgBox("¿Desea Modificar el dato seleccionado?", vbYesNo, "Agregar Aportacion")
    If z = vbYes Then
        ciaportacion = GridAportacion.TextMatrix(GridAportacion.Row, o)
        idaportacion = ModuloFunciones.buscarID("aportaciones", "idaportaciones", ciaportacion) 'se selecciona la tabla de cliente y se realiza la comparacion del compo ciruc con la cedula
        
        Set TBLAportacion = Nothing ' se usa para vaciar la variable recordste y poder agg mas datos
        TBLAportacion.Open "select * from  aportaciones  where idaportaciones ='" & ciaportacion & "'", CONEXION, adOpenDynamic, adLockOptimistic
        TBLAportacion.MoveFirst
        
        Me.TxtAportaciones.Text = TBLAportacion.Fields("aportacion").Value
        Me.TxtValor.Text = TBLAportacion.Fields("valor").Value
        Me.TxtDescripcion.Text = IIf(IsNull(TBLAportacion.Fields("descripcion").Value), "", TBLAportacion.Fields("descripcion").Value)
        'Me.txtRazon_Social.Text = IIf(IsNull(TBLCliente.Fields("razonsocial").Value), "", TBLCliente.Fields("razonsocial").Value)
     
        Else
        
        End If
End Sub
Private Sub TxtAportaciones_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then TxtDescripcion.SetFocus
        KeyAscii = ModuloFunciones.letras(KeyAscii)
End Sub
Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then TxtValor.SetFocus
        KeyAscii = ModuloFunciones.letras(KeyAscii)
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)
        KeyAscii = ModuloFunciones.Numeros(KeyAscii)
End Sub
Private Sub configurargrid()
 With GridAportacion
 .Clear
 .FormatString = "Codigo|Valor|Aportacion|Descripcion"
 .ColWidth(0) = 1000
 .ColWidth(1) = 800
 .ColWidth(2) = 1400
 .ColWidth(3) = 1600
 End With
End Sub
