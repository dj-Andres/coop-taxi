VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAgregarMultas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Multas"
   ClientHeight    =   6240
   ClientLeft      =   9105
   ClientTop       =   3615
   ClientWidth     =   9555
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmAgregarMultas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9555
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
      Left            =   240
      Picture         =   "FrmAgregarMultas.frx":150E1
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1500
   End
   Begin VB.TextBox txtValor 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      MaxLength       =   3
      TabIndex        =   6
      Top             =   2160
      Width           =   2295
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
      Left            =   240
      Picture         =   "FrmAgregarMultas.frx":197D7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
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
      Left            =   240
      Picture         =   "FrmAgregarMultas.frx":1D68D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   7335
      Begin MSFlexGridLib.MSFlexGrid GridMultas 
         Height          =   2055
         Left            =   480
         TabIndex        =   10
         Top             =   2880
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin VB.TextBox txtObservacion 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox TxtMultas 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MaxLength       =   35
         TabIndex        =   1
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label2 
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
         Left            =   600
         TabIndex        =   8
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observacion"
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
         Left            =   600
         TabIndex        =   7
         Top             =   1200
         Width           =   1680
      End
      Begin VB.Label LblMultas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multas"
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
         Left            =   600
         TabIndex        =   2
         Top             =   600
         Width           =   915
      End
   End
End
Attribute VB_Name = "FrmAgregarMultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idmulta As Integer
Dim cimulta As Integer
Dim TBLGuardarMulta As New ADODB.Recordset
Dim TBLMulta As New ADODB.Recordset
Private Sub Cmdguardar_Click()
    If Len(TxtMultas.Text) > 1 And Len(txtObservacion.Text) > 1 And Len(TxtValor.Text) > 1 Then
        Dim respuestaID As Integer
        Set TBLGuardarMulta = Nothing
            TBLGuardarMulta.Open "select * from multas where nombre_multa= '" & UCase(TxtMultas.Text) & "'", CONEXION, adOpenDynamic, adLockOptimistic
            If Not (TBLGuardarMulta.EOF) Then
                MsgBox "La multa ya existe"
            
            
            Else
                'Dim sql As String
                'sql = "insert into aportaciones (aportacion, descripcion, valor)values('" & UCase(TxtAportaciones.Text) & "', '" & UCase(TxtDescripcion.Text) & "', " & UCase(TxtValor.Text) & ")"
                'MsgBox sql
                Set TBLGuardarMulta = Nothing
                TBLGuardarMulta.Open "select * from multas", CONEXION, adOpenDynamic, adLockOptimistic
                TBLGuardarMulta.AddNew
                TBLGuardarMulta!nombre_multa = Trim(UCase(TxtMultas.Text))
                TBLGuardarMulta!observacion = Trim(UCase(Me.txtObservacion.Text))
                TBLGuardarMulta!valor = TxtValor.Text
                
                TBLGuardarMulta.Update
                'Call cargargrid
                
                'Set TBLGuardarAportacion = CONEXION.Execute(sql)
                Call CargarMulta
            End If
    Else
        MsgBox "Ingrese registro", vbInformation, "Agregar Multas"
    End If
    TxtMultas.Text = ""
    txtObservacion.Text = ""
    TxtValor.Text = ""
        
End Sub

Private Sub cmdModificar_Click()
    If Len(TxtMultas.Text) > 1 And Len(txtObservacion.Text) > 1 And Len(TxtValor.Text) > 1 Then
        Dim respuestaID As Integer
        Set TBLGuardarMulta = Nothing
            'TBLGuardarAportacion.Open "select * from aportaciones where aportacion= '" & UCase(TxtAportaciones.Text) & "'", CONEXION, adOpenDynamic, adLockOptimistic
            'If Not (TBLGuardarAportacion.EOF) Then
                'MsgBox "La aportacion ya existe"
            
            
            'Else
                'Dim sql As String
                'sql = "insert into aportaciones (aportacion, descripcion, valor)values('" & UCase(TxtAportaciones.Text) & "', '" & UCase(TxtDescripcion.Text) & "', " & UCase(TxtValor.Text) & ")"
                'MsgBox sql
                Set TBLGuardarMulta = Nothing
                TBLGuardarMulta.Open "select * from multas where idmulta=" & cimulta, CONEXION, adOpenDynamic, adLockOptimistic
                'TBLGuardarAportacion.AddNew
                TBLGuardarMulta!nombre_multa = Trim(UCase(TxtMultas.Text))
                TBLGuardarMulta!observacion = Trim(UCase(Me.txtObservacion.Text))
                TBLGuardarMulta!valor = TxtValor.Text
                
                TBLGuardarMulta.Update
                'Call cargargrid
                
                'Set TBLGuardarAportacion = CONEXION.Execute(sql)
                Call CargarMulta
            End If
    'Else
        'MsgBox "Ingrese registro"
    'End If
    TxtMultas.Text = ""
    txtObservacion.Text = ""
    TxtValor.Text = ""
    
End Sub

Private Sub CmdRegresar_Click()
    Unload Me
    frmGestionMultas.Enabled = True
End Sub
Private Sub Form_Load()
    Call configurargrid
    Call CargarMulta
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmGestionMultas.Enabled = True
End Sub
Private Sub CargarMulta()
    Dim sql As String
    sql = "select * from multas"
    Set TBLMultas = CONEXION.Execute(sql)
    Dim f As Integer
    f = 1
    GridMultas.Rows = 2
    Do Until TBLMultas.EOF
        GridMultas.TextMatrix(f, 0) = TBLMultas!idmulta
        GridMultas.TextMatrix(f, 1) = TBLMultas!nombre_multa
        GridMultas.TextMatrix(f, 2) = IIf(IsNull(TBLMultas!observacion), "", TBLMultas!observacion)
        GridMultas.TextMatrix(f, 3) = TBLMultas!valor
        TBLMultas.MoveNext
        f = f + 1
        GridMultas.Rows = GridMultas.Rows + 1
    Loop


End Sub
Private Sub GridMultas_DblClick()
    Dim z As Integer
    z = MsgBox("¿Desea Modificar el dato seleccionado?", vbYesNo, "Agregar Multas")
    If z = vbYes Then
        cimulta = GridMultas.TextMatrix(GridMultas.Row, o)
        idmulta = ModuloFunciones.buscarID("multas", "idmulta", cimulta) 'se selecciona la tabla de cliente y se realiza la comparacion del compo ciruc con la cedula
        
        Set TBLMulta = Nothing ' se usa para vaciar la variable recordste y poder agg mas datos
        TBLMulta.Open "select * from  multas  where idmulta ='" & cimulta & "'", CONEXION, adOpenDynamic, adLockOptimistic
        TBLMulta.MoveFirst
        
        Me.TxtMultas.Text = TBLMulta.Fields("nombre_multa").Value
        Me.txtObservacion.Text = IIf(IsNull(TBLMulta.Fields("observacion").Value), "", TBLMulta.Fields("observacion").Value)
        Me.TxtValor.Text = TBLMulta.Fields("valor").Value
        'Me.txtRazon_Social.Text = IIf(IsNull(TBLCliente.Fields("razonsocial").Value), "", TBLCliente.Fields("razonsocial").Value)
     
        Else
        
        End If
End Sub

Private Sub TxtMultas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtObservacion.SetFocus
        KeyAscii = ModuloFunciones.letras(KeyAscii)
End Sub
Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtValor.SetFocus
        KeyAscii = ModuloFunciones.letras(KeyAscii)
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = ModuloFunciones.Numeros(KeyAscii)
End Sub
Private Sub configurargrid()
     With GridMultas
 .Clear
 .FormatString = "Codigo|Multa|Observacion|Valor"
 .ColWidth(0) = 600
 .ColWidth(1) = 1900
 .ColWidth(2) = 1900
 .ColWidth(3) = 700
 End With
End Sub
