VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FCambiousuario 
   Caption         =   "Cambio de Usuarios"
   ClientHeight    =   2730
   ClientLeft      =   5730
   ClientTop       =   2430
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   3975
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   2280
      OleObjectBlob   =   "FCambiousuario.frx":0000
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FCambiousuario.frx":0084
      Height          =   1335
      Left            =   5160
      TabIndex        =   5
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2355
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
   Begin MSAdodcLib.Adodc Adodc_usuario 
      Height          =   330
      Left            =   5160
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sistema Ticket\DBS\DBsistemaTicket.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sistema Ticket\DBS\DBsistemaTicket.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM Tacceso"
      Caption         =   "Adodc1"
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FCambiousuario.frx":00A0
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FCambiousuario.frx":0112
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton CBaceptar 
      Caption         =   "Guardar Nuevo Usuario"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox CTusuario 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1080
      MaxLength       =   9
      TabIndex        =   0
      Text            =   "Nuevo"
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox CTcontraseña 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "Nueva"
      Top             =   1200
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FCambiousuario.frx":017E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   2280
      OleObjectBlob   =   "FCambiousuario.frx":03B2
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "FCambiousuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CBaceptar_Click()
Adodc_usuario.RecordSource = "select * from Tacceso"
Adodc_usuario.Refresh

Adodc_usuario.Recordset.MoveFirst
        Adodc_usuario.Recordset.Fields(0).Value = CTusuario.Text
        Adodc_usuario.Recordset.Fields(1).Value = CTcontraseña.Text
        Adodc_usuario.Recordset.Update
        Adodc_usuario.Refresh
        Adodc_usuario.Refresh
MsgBox "Nuevo Usuario: " & CTusuario.Text, vbInformation, "Favor de Anotar nuevo Usuario"
MsgBox "Nueva Contraseña: " & CTcontraseña.Text, vbInformation, "Favor de Anotar nueva Contraseña"
Unload Me
End Sub

Private Sub CTcontraseña_Change()
Call Fnormalidad
    Call Fefecto("CTcontraseña")
End Sub

Private Sub CTusuario_Change()
 Call Fnormalidad
    Call Fefecto("CTusuario")
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin ("C:\Sistema Ticket\SKN\vector cell.skn")
    Skin1.ApplySkin Me.hWnd
    Call Flimpiar
    Call Fnormalidad
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


