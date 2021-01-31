VERSION 5.00
Begin VB.Form frmIngreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6540
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmIngreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   5295
      Begin VB.TextBox txTClave 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtUsuario 
         Height          =   495
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   5
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Clave"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdContinuar 
      Caption         =   "Continuar"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      Picture         =   "frmIngreso.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "Regresar"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Picture         =   "frmIngreso.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   2
      Top             =   -240
      Width           =   1755
   End
End
Attribute VB_Name = "frmIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tblinicio As New ADODB.Recordset
Private Sub cmdContinuar_Click()

Set tblinicio = CONEXION.Execute("select * from usuarios")
        If Me.txtUsuario.Text = tblinicio!usuario And Me.txTClave.Text = tblinicio!clave Then
        
        frmMenu.Show
        Else
MsgBox "EL USUARIO O LA CONTRASEÑA ES INCORRECTA VUELVA A INTENTARLO"
Me.txtUsuario = ""
Me.txTClave = ""
End If


            Unload Me
'End If




End Sub

Private Sub CmdRegresar_Click()
frmLogin.Show
frmIngreso.Hide

End Sub

Private Sub Form_Load()
ModuloBaseDatos.conectardb
End Sub

Private Sub txTClave_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Me.cmdContinuar.SetFocus
        KeyAscii = ModuloFunciones.Direccion(KeyAscii)
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Me.txTClave.SetFocus
        KeyAscii = ModuloFunciones.letras(KeyAscii)
End Sub
