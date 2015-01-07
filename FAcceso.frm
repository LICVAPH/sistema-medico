VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FAcceso 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acceso al Sistema"
   ClientHeight    =   1995
   ClientLeft      =   9360
   ClientTop       =   5490
   ClientWidth     =   5580
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5580
   Begin VB.CommandButton Command1 
      Height          =   1575
      Left            =   120
      Picture         =   "FAcceso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   1680
      OleObjectBlob   =   "FAcceso.frx":1131
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   1800
      OleObjectBlob   =   "FAcceso.frx":11A3
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "FAcceso.frx":1211
      Top             =   2400
   End
   Begin VB.CommandButton Bayuda 
      Caption         =   "..."
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Busuarios 
      Caption         =   "Accesos"
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Bintrusos 
      Caption         =   "Intrusos"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FAcceso.frx":1445
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   6840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "usuario"
         Caption         =   "usuario"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "contraseña"
         Caption         =   "contraseña"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Bingreso 
      Caption         =   "Ingresar al Sistema"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox CTcontraseña 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   9
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox CTusuario 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2640
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Frame Frame_accesos 
      Caption         =   "[ Catalogo de Accesos ]"
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   6015
      Begin MSAdodcLib.Adodc Adodc_usuario 
         Height          =   375
         Left            =   120
         Top             =   2640
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sistema Medico\DBS\DBsistemaMedico.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sistema Medico\DBS\DBsistemaMedico.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from Tcatalogo_usuarios"
         Caption         =   "Usuarios"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "FAcceso.frx":1460
         Height          =   1695
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2990
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "ID"
            Caption         =   "ID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "USUARIO"
            Caption         =   "USUARIO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "FECHA"
            Caption         =   "FECHA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "HORA"
            Caption         =   "HORA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc_acceso 
      Height          =   375
      Left            =   4680
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sistema Medico\DBS\DBsistemaMedico.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sistema Medico\DBS\DBsistemaMedico.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Tacceso"
      Caption         =   "Acceso"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame_intrusos 
      Caption         =   "[ Catalogo de Intrusos ]"
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   6015
      Begin MSAdodcLib.Adodc Adodc_intruso 
         Height          =   375
         Left            =   120
         Top             =   2640
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sistema Medico\DBS\DBsistemaMedico.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sistema Medico\DBS\DBsistemaMedico.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from Tcatalogo_intrusos"
         Caption         =   "Intrusos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "FAcceso.frx":147C
         Height          =   1695
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2990
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "ID"
            Caption         =   "ID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "USUARIO"
            Caption         =   "USUARIO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "FECHA"
            Caption         =   "FECHA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "HORA"
            Caption         =   "HORA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bayuda_Click()
Static Vestado As Integer
If Vestado = 1 Then
    Me.Height = 2535
    Me.Width = 5800
    Vestado = 0
ElseIf Vestado = 0 Then
    Me.Height = 6960
    Me.Width = 6375
    Vestado = 1
End If
End Sub

Private Sub Bingreso_Click()
Static VAR As Integer
Adodc_acceso.RecordSource = "SELECT * FROM Tacceso WHERE USUARIO='" + CTusuario.Text + "' AND CONTRASEÑA='" + CTcontraseña.Text + "'"
Adodc_acceso.Refresh
If Adodc_acceso.Recordset.EOF = True Then
    VAR = VAR + 1
    If VAR = 3 Then
        MsgBox " Se Ha Detectado un Intruso Y el Sistema Deve Cerrarse"
        Call Fintrusos
        Unload Me
    Else
       Call Fintrusos
       Call Flimpiar
       Adodc_usuario.RecordSource = "select * from tcatalogo_usuarios"
       Adodc_usuario.Refresh
       Adodc_usuario.Refresh
       
       Adodc_intruso.RecordSource = "select * from tcatalogo_intrusos"
       Adodc_intruso.Refresh
       Adodc_intruso.Refresh
    End If
Else
    Adodc_acceso.RecordSource = "SELECT * FROM Tcadena"
    Adodc_acceso.Refresh
    If Adodc_acceso.Recordset.RecordCount = 0 Then
        Call Fusuarios
        Unload Me
        Asignacion.Show
    Else
        Call Fusuarios
        Unload Me
        FMenu.Show
    End If
End If
End Sub
Private Sub Bintrusos_Click()
    Frame_accesos.Visible = False
    Frame_intrusos.Visible = True
End Sub
Private Sub Busuarios_Click()
    Frame_accesos.Visible = True
    Frame_intrusos.Visible = False
End Sub
Private Sub CTcontraseña_Change()
    Call Fnormalidad
    Call Fefecto("CTcontraseña")
End Sub
Private Sub CTcontraseña_GotFocus()
    Call Fnormalidad
End Sub
Private Sub CTcontraseña_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Bingreso.SetFocus
    End If
End Sub
Private Sub CTusuario_Change()
    Call Fnormalidad
    Call Fefecto("CTusuario")
End Sub
Private Sub CTusuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CTcontraseña.SetFocus
    End If
End Sub
Private Sub CTusuario_GotFocus()
    Call Fnormalidad
End Sub
Private Sub Form_Load()
    Skin1.LoadSkin ("C:\Sistema Ticket\SKN\vector cell.skn")
    Skin1.ApplySkin Me.hWnd
    Call Flimpiar
    Call Fnormalidad
    Me.Caption = MPropiedades.FMdatosdelsistema
    
End Sub
'''''''''''seccion de funciones''''''''''''''''''''''''''''''''''''

'funcion que limpia todas las cajas de texto
Public Function Flimpiar()
Dim Vcomponente As Control
    For Each Vcomponente In Me.Controls
   If TypeOf Vcomponente Is TextBox Then
       Vcomponente.Text = Empty
    Else
        'no pasa nada
    End If
Next
End Function '''''''''''''''termina funcion
'funcion que regresa las cajas a su normalidad
Public Function Fnormalidad()
Dim Vcomponente As Control
    For Each Vcomponente In Me.Controls
   If TypeOf Vcomponente Is TextBox Then
        Vcomponente.FontSize = 10
        Vcomponente.FontBold = False
        Vcomponente.Height = 375
        Vcomponente.ForeColor = &H0&
         Vcomponente.BackColor = &H80000005
    Else
        'no pasa nada
    End If
Next
End Function ''''''''''''''''''termina funcion
'funcion que da efecto a las cajas de texto
Public Function Fefecto(Vnomobjeto As String)
Dim Vcomponente As Control
For Each Vcomponente In Me.Controls
        If TypeOf Vcomponente Is TextBox Then
            If Vnomobjeto = Vcomponente.Name Then
            Vcomponente.FontSize = 16
            Vcomponente.FontBold = True
            Vcomponente.Height = 425
            Vcomponente.ForeColor = &H0&
            Vcomponente.BackColor = &HDC9367
            Else
                'no pasa nada
            End If
        Else
            'no pasa nada
        End If
Next
End Function ''''''''''''''''''termina funcion

'funcion que almacena los ingresos al sistema
Public Function Fusuarios()
        Adodc_usuario.Recordset.AddNew
        Adodc_usuario.Recordset.Fields(1).Value = CTusuario.Text
        Adodc_usuario.Recordset.Fields(2).Value = Date
        Adodc_usuario.Recordset.Fields(3).Value = Time
        Adodc_usuario.Recordset.Update
        Adodc_usuario.Refresh
        Adodc_usuario.Refresh
End Function '''''funcion que almacena los ingresos al sistema

'funcion que almacena los intrusos o errores al sistema
Public Function Fintrusos()
        Adodc_intruso.Recordset.AddNew
        Adodc_intruso.Recordset.Fields(1).Value = CTusuario.Text
        Adodc_intruso.Recordset.Fields(2).Value = Date
        Adodc_intruso.Recordset.Fields(3).Value = Time
        Adodc_intruso.Recordset.Update
        Adodc_intruso.Refresh
        Adodc_intruso.Refresh
End Function '''funcion que almacena los intrusos o errores al sistema

'''''''''''termina seccion de funciones''''''''''''''''''''''''''''




Private Sub TabStrip1_Change()

End Sub

Private Sub TabStrip1_Click()

End Sub
