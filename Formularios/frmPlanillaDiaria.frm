VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlanillaDiaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla Diaria"
   ClientHeight    =   7875
   ClientLeft      =   300
   ClientTop       =   1365
   ClientWidth     =   14475
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   14475
   Begin VB.TextBox txtVariantes 
      Height          =   315
      Left            =   1200
      TabIndex        =   38
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox txtTotalRemis 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   450
      Left            =   10680
      TabIndex        =   1
      Top             =   7410
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   13335
      TabIndex        =   3
      Top             =   7410
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Cancelar"
      Height          =   450
      Left            =   12450
      TabIndex        =   2
      Top             =   7410
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7320
      Left            =   15
      TabIndex        =   11
      Top             =   30
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   12912
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   512
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "frmPlanillaDiaria.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameFactura"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmPlanillaDiaria.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "frameBuscar"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5805
         Left            =   105
         TabIndex        =   20
         Top             =   1455
         Width           =   14235
         Begin VB.CommandButton cmdAgregarProducto 
            Height          =   330
            Left            =   13800
            MaskColor       =   &H8000000F&
            Picture         =   "frmPlanillaDiaria.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Agregar Cliente"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   405
            Left            =   4095
            MaxLength       =   60
            TabIndex        =   22
            Top             =   5340
            Width           =   10050
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   5340
            Width           =   1215
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   300
            TabIndex        =   23
            Top             =   525
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   5130
            Left            =   75
            TabIndex        =   24
            Top             =   165
            Width           =   13710
            _ExtentX        =   24183
            _ExtentY        =   9049
            _Version        =   393216
            Rows            =   3
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   290
            BackColorSel    =   12648447
            ForeColorSel    =   0
            GridColor       =   -2147483633
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            ScrollBars      =   2
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdQuitarProducto 
            Height          =   330
            Left            =   13800
            MaskColor       =   &H8000000F&
            Picture         =   "frmPlanillaDiaria.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Quitar Cliente"
            Top             =   720
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2730
            TabIndex        =   25
            Top             =   5475
            Width           =   1290
         End
      End
      Begin VB.Frame frameBuscar 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Left            =   -74805
         TabIndex        =   14
         Top             =   420
         Width           =   14070
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   2520
            TabIndex        =   6
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   40763393
            CurrentDate     =   43174
         End
         Begin VB.TextBox txtBuscarCliDescri 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3180
            MaxLength       =   50
            TabIndex        =   5
            Tag             =   "Descripción"
            Top             =   330
            Width           =   4155
         End
         Begin VB.TextBox txtBuscaCliente 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2490
            MaxLength       =   40
            TabIndex        =   4
            Top             =   330
            Width           =   675
         End
         Begin VB.ComboBox cboFactura1 
            Height          =   315
            Left            =   2490
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   990
            Width           =   2400
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Buscar"
            Height          =   390
            Left            =   10200
            MaskColor       =   &H000000FF&
            TabIndex        =   9
            ToolTipText     =   "Buscar "
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   6240
            TabIndex        =   7
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   40763393
            CurrentDate     =   43174
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   1395
            TabIndex        =   19
            Top             =   375
            Width           =   555
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Factura:"
            Height          =   195
            Left            =   1395
            TabIndex        =   18
            Top             =   1020
            Width           =   960
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   5130
            TabIndex        =   16
            Top             =   720
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1395
            TabIndex        =   15
            Top             =   720
            Width           =   990
         End
      End
      Begin VB.Frame FrameFactura 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   105
         TabIndex        =   13
         Top             =   345
         Width           =   14040
         Begin VB.TextBox txtMenu 
            Height          =   315
            Left            =   1080
            TabIndex        =   37
            Top             =   240
            Width           =   4335
         End
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   11520
            TabIndex        =   32
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   40763393
            CurrentDate     =   43169
         End
         Begin VB.ComboBox cboVariantes 
            Height          =   315
            ItemData        =   "frmPlanillaDiaria.frx":10C4
            Left            =   6810
            List            =   "frmPlanillaDiaria.frx":10C6
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   600
            Visible         =   0   'False
            Width           =   4000
         End
         Begin VB.ComboBox cboMenu 
            Height          =   315
            ItemData        =   "frmPlanillaDiaria.frx":10C8
            Left            =   6810
            List            =   "frmPlanillaDiaria.frx":10CA
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   240
            Visible         =   0   'False
            Width           =   4000
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar"
            Height          =   615
            Left            =   12840
            TabIndex        =   31
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblfecha 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   5880
            TabIndex        =   30
            Top             =   480
            Width           =   4455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Variantes:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Menu:"
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   27
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   10920
            TabIndex        =   17
            Top             =   540
            Width           =   495
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   5145
         Left            =   -74835
         TabIndex        =   10
         Top             =   1935
         Width           =   14100
         _ExtentX        =   24871
         _ExtentY        =   9075
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorSel    =   16761024
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   12
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   11565
      TabIndex        =   0
      Top             =   7410
      Width           =   870
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   7080
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   360
      TabIndex        =   33
      Top             =   7560
      Width           =   2460
   End
End
Attribute VB_Name = "frmPlanillaDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim precio_comida, precio_sopa, precio_postre, precio_pan, precio_descartable, precio_remise, monto_total As Double
Dim monto_sopa, monto_postre, monto_pan, monto_descartable, monto_remise, vianda As Double

Private Sub cmdAgregarProducto_Click()
      BuscarClientes
End Sub
Public Sub BuscarClientes()
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Set B = New CBusqueda
    With B
        cSQL = "SELECT CLI_RAZSOC, CLI_CODIGO,CLI_NRODOC"
        cSQL = cSQL & " FROM CLIENTE C"
        
        hSQL = "Nombre, Código, DNI"
        .sql = cSQL
        .Headers = hSQL
        .Field = "CLI_RAZSOC"
        campo1 = .Field
        .Field = "CLI_CODIGO"
        campo2 = .Field
        .Field = "CLI_NRODOC"
        campo3 = .Field
        
        .OrderBy = "CLI_RAZSOC"
        camponumerico = False
        .Titulo = "Busqueda de Clientes :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
'        If .ResultFields.Count > 0 Then
'            If Txt = "txtcodCli" Then
'                txtcodigo.Text = .ResultFields(2)
'                'txtCodCli_LostFocus
'            Else
'                If .ResultFields(3) = "" Then
'                    txtBuscaCliente.Text = .ResultFields(2)
'                    txtcodigo.Text = .ResultFields(2)
'                Else
'                    txtBuscaCliente.Text = .ResultFields(3)
'                End If
'                txtBuscaCliente_LostFocus
'            End If
'        End If
    End With

    Set B = Nothing
End Sub

Private Sub CmdBuscar_Click()
    grdGrilla.Rows = 1
    
    'Verifico si ya fue creada la planilla
    sql = "SELECT PDI_FECHA FROM PLANILLA_DIARIA "
    sql = sql & " WHERE PDI_FECHA=" & XDQ(Fecha)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = True Then
        'NUEVA PLANILLA
        crearplanilla
        lblEstado.Caption = "Planilla Nueva"
        cmdImprimir.Enabled = False
    Else
        'PLANILLA EXISTENTE
        cargarplanilla
        lblEstado.Caption = "Planilla Existente"
        cmdImprimir.Enabled = True
    End If
    Rec1.Close
    
    
    
End Sub
'Private Function preciovianda() As Double
'    preciovianda = 0
'    Rec1.Open "SELECT VIA_PRECIO FROM VIANDAS WHERE VIA_CODIGO=1", DBConn, adOpenStatic, adLockOptimistic
'    If Rec1.EOF = False Then
'        preciovianda = Format(Rec1!VIA_PRECIO, "#,##0.00")
'    End If
'    Rec1.Close
'
'End Function
Private Sub cmdGrabar_Click()
   'Verifico si ya fue creada la planilla
   sql = "SELECT PDI_FECHA FROM PLANILLA_DIARIA "
   sql = sql & " WHERE PDI_FECHA=" & XDQ(Fecha)
    
   rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
   If rec.EOF = True Then
       If MsgBox("Confirma la planilla diaria del: " & lblfecha.Caption, vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            'PLANILLA_DIARIA
            sql = " INSERT INTO PLANILLA_DIARIA"
            sql = sql & "(PDI_FECHA,PDI_PRINCI, PDI_VARIAN, PDI_TOTAL,PDI_TOTREM, PDI_OBSERVA)"
            sql = sql & " VALUES ("
            sql = sql & XDQ(Fecha.Value) & ","
            'sql = sql & cboMenu.ItemData(cboMenu.ListIndex) & ","
            'sql = sql & cboVariantes.ItemData(cboVariantes.ListIndex) & ","
            sql = sql & XS(txtMenu.Text) & ","
            sql = sql & XS(txtVariantes.Text) & ","
            sql = sql & XN(txtTotal.Text) & ","
            sql = sql & XN(txtTotalRemis.Text) & ","
            sql = sql & XS(txtObservaciones.Text) & ")"
            DBConn.Execute sql
            
            'PLANILLA_DIARIA_DETALLE
            For i = 1 To grdGrilla.Rows - 1
                'INSERTAR NUEVOS
                sql = " INSERT INTO PLANILLA_DIARIA_DETALLE "
                sql = sql & "(PDI_FECHA,CLI_CODIGO, PDI_PRECIO, PDI_ALMUER, PDI_CENA,"
                sql = sql & "PDI_SOPA,PDI_POSTRE,PDI_PAN,PDI_DESCAR,PDI_REMISE)"
                sql = sql & " VALUES ("
                sql = sql & XDQ(Fecha.Value) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 14)) & "," 'CODIGO_CLIENTE
                sql = sql & XN(grdGrilla.TextMatrix(i, 0)) & "," 'PDI_PRECIO
                sql = sql & XN(grdGrilla.TextMatrix(i, 3)) & "," 'PDI_ALMUER
                sql = sql & XN(grdGrilla.TextMatrix(i, 4)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 5)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 6)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 7)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 8)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 16)) & ")"
                DBConn.Execute sql
            Next
       End If
       lblEstado.Caption = "Planilla Existente"
   Else
        If MsgBox("La planilla diaria del dia: " & lblfecha.Caption & " ya ha sido grabada. " & Chr(13) & "Desea modificarla?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            'ACTUALIZO PLANILLA_DIARIA
            
            'PLANILLA_DIARIA
            sql = " UPDATE PLANILLA_DIARIA SET "
            sql = sql & "PDI_PRINCI=" & XS(txtMenu.Text)
            sql = sql & ",PDI_VARIAN=" & XS(txtVariantes.Text)
            sql = sql & ",PDI_TOTAL=" & XN(txtTotal.Text)
            sql = sql & ",PDI_OBSERVA=" & XS(txtObservaciones.Text)
            sql = sql & " WHERE PDI_FECHA = " & XDQ(Fecha)
            DBConn.Execute sql
            
            
            sql = "DELETE FROM PLANILLA_DIARIA_DETALLE WHERE PDI_FECHA=" & XDQ(Fecha)
            DBConn.Execute sql
            
            For i = 1 To grdGrilla.Rows - 1
                'INSERTAR NUEVOS
                sql = " INSERT INTO PLANILLA_DIARIA_DETALLE "
                sql = sql & "(PDI_FECHA,CLI_CODIGO, PDI_PRECIO, PDI_ALMUER, PDI_CENA,"
                sql = sql & "PDI_SOPA,PDI_POSTRE,PDI_PAN,PDI_DESCAR,PDI_REMISE)"
                sql = sql & " VALUES ("
                sql = sql & XDQ(Fecha.Value) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 14)) & "," 'CODIGO_CLIENTE
                sql = sql & XN(grdGrilla.TextMatrix(i, 0)) & "," 'PDI_PRECIO
                sql = sql & XN(grdGrilla.TextMatrix(i, 3)) & "," 'PDI_ALMUER
                sql = sql & XN(grdGrilla.TextMatrix(i, 4)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 5)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 6)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 7)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 8)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 16)) & ")"
                DBConn.Execute sql
              Next
        End If
        lblEstado.Caption = "Planilla Existente"
    End If
    rec.Close
    
    If MsgBox("Desea imprimir la PLANILLA DE INDICACIONES Y DISTRIBUCION del dia: " & lblfecha.Caption, vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        cmdImprimir_Click
    End If
End Sub

Private Sub cmdImprimir_Click()
    Screen.MousePointer = vbHourglass
  
    
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.SortFields(0) = ""
    
    Rep.SelectionFormula = " {PLANILLA_DIARIA.PDI_FECHA} >= DATE (" & Mid(Fecha.Value, 7, 4) & "," & Mid(Fecha.Value, 4, 2) & "," & Mid(Fecha.Value, 1, 2) & ")"
    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    
    Rep.WindowTitle = "Planilla de Indicacion y Distribucion"
    Rep.ReportFileName = DirReport & "planilla_diaria.rpt"
    Rep.Action = 1
'    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    Rep.SelectionFormula = ""
End Sub
Private Function limpiarplanilla()
    lblEstado.Caption = "Planilla Nueva"
    grdGrilla.Rows = 1
    txtVariantes.Text = ""
    txtMenu.Text = ""
    
    cmdImprimir.Enabled = False
    txtTotal.Text = "0,00"
    txtTotalRemis.Text = "0,00"
End Function
Private Sub CmdNuevo_Click()
    limpiarplanilla
    Fecha = Date
    Fecha_LostFocus
End Sub

Private Sub cmdQuitarProducto_Click()
    If MsgBox("Seguro que desea quitar el Cliente: " & grdGrilla.TextMatrix(grdGrilla.RowSel, 2), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        If grdGrilla.Rows > 2 Then
            grdGrilla.RemoveItem (grdGrilla.RowSel)
        Else
            grdGrilla.Rows = 1
        End If
        txtTotal = SumaTotal
        txtTotalRemis = SumaTotalRemis
        txtTotal = Valido_Importe(txtTotal)
    End If
End Sub

Private Sub cmdReporte_Click()
    
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmPlanillaDiaria = Nothing
        Unload Me
    End If
End Sub

Private Sub Fecha_Change()
    lblfecha.Caption = FechaLetras(Fecha.day, Fecha.dayofweek, Fecha.month, Fecha.year)
    If lblEstado.Caption = "Planilla Existente" Then
        limpiarplanilla
    End If
End Sub

Private Sub Fecha_LostFocus()
    lblfecha.Caption = FechaLetras(Fecha.day, Fecha.dayofweek, Fecha.month, Fecha.year)
End Sub

Private Function crearplanilla()
    'obtener que dia es hoy, segun fecha actual y programa
    'verificar si es feriado
    'buscar clientes por programa_cliente
    Dim DIA As Integer
    Dim Cli As Integer
    Dim i As Integer
    DIA = dia_programa(Fecha)
    sql = "SELECT  * FROM CLIENTE C, PROGRAMA_CLIENTE PC, LOCALIDAD L"
    sql = sql & " WHERE C.CLI_CODIGO=PC.CLI_CODIGO "
    sql = sql & " AND C.LOC_CODIGO = L.LOC_CODIGO "
    sql = sql & " AND (PC.PRG_CODIGO = " & DIA & " OR PC.PRG_CODIGO= " & DIA + 8 & ")" 'ALMUERZO MAS CENA DE ESE
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Cli = 0
        i = 0
        Do While Not rec.EOF
            If rec!CLI_CODIGO = Cli Then
                'ACTUALIZAR COLUMMNA CENA
                grdGrilla.TextMatrix(i, 4) = rec!PRC_CANT
            Else
                Cli = rec!CLI_CODIGO
                grdGrilla.AddItem Chk0(rec!CLI_TOTAL) & Chr(9) & _
                                  TIPO_FAC(rec!CLI_FACTURA) & Chr(9) & _
                                  rec!CLI_RAZSOC & Chr(9) & _
                                  rec!PRC_CANT & Chr(9) & _
                                   "" & Chr(9) & _
                                   "" & Chr(9) & _
                                   "" & Chr(9) & _
                                   "" & Chr(9) & _
                                   "" & Chr(9) & _
                                   "" & Chr(9) & _
                                   "" & Chr(9) & _
                                   rec!CLI_DIAGNO & Chr(9) & _
                                   rec!LOC_DESCRI & Chr(9) & _
                                   "" & Chr(9) & _
                                   rec!CLI_CODIGO
                                   

                
                i = i + 1
            End If
            rec.MoveNext
        Loop
    End If
    rec.Close
    
    'SACAR DE CLIENTE VIANDAS SI TIENE POSTRE, PAN, DESCARTABLES O REMISE
    ' SI EL CLIENTE TIENE 2 ALMUERZOS Y TIENE SELECCIONADO POSTRE
    ' PAN Y/O DESCARTABLES, PONER UN DOS. O PONER UN SI/NO
    'Recorro la grilla y cargo lo q se incluye de la vianda
    If grdGrilla.Rows > 1 Then
        For i = 1 To grdGrilla.Rows - 1
            
            sql = "SELECT  * FROM CLIENTE C, CLIENTE_VIANDAS CV, VIANDAS V"
            sql = sql & " WHERE C.CLI_CODIGO=CV.CLI_CODIGO "
            sql = sql & " AND CV.VIA_CODIGO=V.VIA_CODIGO"
            sql = sql & " AND C.CLI_CODIGO=" & grdGrilla.TextMatrix(i, 14)
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                Do While rec.EOF = False
                    Select Case rec!VIA_CODIGO
                        Case 2 ' Sopa
                            grdGrilla.TextMatrix(i, 5) = 1
                        Case 3 ' Postre
                            grdGrilla.TextMatrix(i, 6) = 1
                        Case 4 ' Pan
                            grdGrilla.TextMatrix(i, 7) = 1
                        Case 5 ' Descartable
                            grdGrilla.TextMatrix(i, 8) = 1
                        Case 6 ' Remise Pilar
                            grdGrilla.TextMatrix(i, 9) = "$ " & Chk0(rec!VIA_PRECIO)
                            grdGrilla.TextMatrix(i, 16) = Chk0(rec!VIA_PRECIO)
                        Case 7 ' Remise Rio 2
                            grdGrilla.TextMatrix(i, 9) = "$ " & Chk0(rec!VIA_PRECIO)
                            grdGrilla.TextMatrix(i, 16) = Chk0(rec!VIA_PRECIO)
                    End Select
                    rec.MoveNext
                Loop
            End If
            'actualizo el montos
            grdGrilla.TextMatrix(i, 15) = grdGrilla.TextMatrix(i, 0) 'precio unitario de comida por defecto
            grdGrilla.TextMatrix(i, 0) = monto(i) ' monto por cantidad de almuerzo/cena
            
            rec.Close
        Next
        txtTotal = SumaTotal
        txtTotalRemis = SumaTotalRemis
        txtTotal = Valido_Importe(txtTotal)
    End If
    
    
End Function

Private Function cargarplanilla()
    'obtener que dia es hoy, segun fecha actual y programa
    'verificar si es feriado
    'buscar clientes por programa_cliente
    Dim DIA As Integer
    Dim Cli As Integer
    Dim i As Integer
    DIA = dia_programa(Fecha)
    sql = "SELECT  *"
    sql = sql & " FROM PLANILLA_DIARIA P"
    sql = sql & " WHERE "
    sql = sql & " PDI_FECHA= " & XDQ(Fecha.Value)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtMenu.Text = ChkNull(rec!PDI_PRINCI)
        txtVariantes.Text = ChkNull(rec!PDI_VARIAN)
        txtTotal.Text = Chk0(rec!PDI_TOTAL)
        txtTotalRemis.Text = Chk0(rec!PDI_TOTREM)
        txtObservaciones.Text = ChkNull(rec!PDI_OBSERVA)
    End If
    rec.Close
    
    sql = "SELECT  P.* ,CL.CLI_RAZSOC,CL.CLI_FACTURA,CL.CLI_DIAGNO, L.LOC_DESCRI"
    sql = sql & " FROM PLANILLA_DIARIA_DETALLE P, CLIENTE CL, LOCALIDAD L"
    sql = sql & " WHERE CL.CLI_CODIGO=P.CLI_CODIGO"
    sql = sql & " AND CL.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND P.PDI_FECHA= " & XDQ(Fecha.Value)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While Not rec.EOF
            grdGrilla.AddItem Chk0(rec!PDI_PRECIO) & Chr(9) & _
                              TIPO_FAC(rec!CLI_FACTURA) & Chr(9) & _
                              rec!CLI_RAZSOC & Chr(9) & _
                              ChkNull(rec!PDI_ALMUER) & Chr(9) & _
                              ChkNull(rec!PDI_CENA) & Chr(9) & _
                              ChkNull(rec!PDI_SOPA) & Chr(9) & _
                              ChkNull(rec!PDI_POSTRE) & Chr(9) & _
                              ChkNull(rec!PDI_PAN) & Chr(9) & _
                              ChkNull(rec!PDI_DESCAR) & Chr(9) & _
                              "$ " & ChkNull(rec!PDI_REMISE) & Chr(9) & _
                              ChkNull(rec!PDI_VIANDA) & Chr(9) & _
                              ChkNull(rec!CLI_DIAGNO) & Chr(9) & _
                              ChkNull(rec!LOC_DESCRI) & Chr(9) & _
                              ChkNull(rec!PDI_OBSERV) & Chr(9) & _
                              ChkNull(rec!CLI_CODIGO) & Chr(9) & _
                              Chk0(rec!PDI_PRECIO) & Chr(9) & _
                              ChkNull(rec!PDI_REMISE)

            i = i + 1
            rec.MoveNext
        Loop
    End If
    rec.Close
    
'    txtTotal = SumaTotal
'    txtTotalRemis = SumaTotalRemis
'    txtTotal = Valido_Importe(txtTotal)
    
End Function

Private Function TIPO_FAC(Codigo As Integer) As String
    Dim letra As String
    letra = ""
    Select Case Codigo
        Case 1 ' DIARIA
            letra = "D"
        Case 2 ' SEMANAL
            letra = "S"
        Case 3 'QUINCENAL
            letra = "Q"
        Case 4 'MENSUAL
            letra = "M"
    End Select
    
    TIPO_FAC = letra
End Function
Private Function monto(Fila As Integer) As Double
'    grdGrilla.ColWidth(3) = 700 'Almuerzo
'    grdGrilla.ColWidth(4) = 700  'Cena
'    grdGrilla.ColWidth(5) = 700  'Sopa
'    grdGrilla.ColWidth(6) = 700  'Postre
'    grdGrilla.ColWidth(7) = 700  'Pan
'    grdGrilla.ColWidth(8) = 700 'Descart.
'    grdGrilla.ColWidth(9) = 700  'Remise
    
    Dim montoaux As Double
    
    montoaux = 0
    'Calcular monto de comida mas, lo que sume en la grilla, por ej, comid 50, sin son 2
    'precio_sopa, precio_postre, precio_pan, precio_descartable, precio_remise
    
    ' precio comida por almuerzo o cenas
    montoaux = precio_comida * (CInt(Chk0(grdGrilla.TextMatrix(Fila, 3))) + CInt(Chk0(grdGrilla.TextMatrix(Fila, 4))))
    
    'sumo precio sopa
    montoaux = montoaux + precio_sopa * (CInt(Chk0(grdGrilla.TextMatrix(Fila, 5))))
    
    'sumo precio postre
    montoaux = montoaux + precio_postre * (CInt(Chk0(grdGrilla.TextMatrix(Fila, 6))))
    
    'sumo precio pan
    montoaux = montoaux + precio_pan * (CInt(Chk0(grdGrilla.TextMatrix(Fila, 7))))
    
    'sumo precio descartable
    montoaux = montoaux + precio_descartable * (CInt(Chk0(grdGrilla.TextMatrix(Fila, 8))))
    
    'sumo precio remise
   ' montoaux = montoaux + precio_remise * (CInt(Chk0(grdGrilla.TextMatrix(Fila, 9))))
    monto = montoaux
End Function
Private Function dia_programa(Fecha As Date) As Integer
    Dim day As String
    day = Weekday(Fecha)
    dia_programa = day
End Function

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
        
    Me.Left = 0
    Me.Top = 0
    preparargrillas
    Fecha.Value = Date
    lblfecha.Caption = FechaLetras(Fecha.day, Fecha.dayofweek, Fecha.month, Fecha.year)
    cargar_comidas
    cargo_precios
    'lblEstado.Caption = ""
    
End Sub
Private Function cargo_precios()
    sql = "SELECT * FROM VIANDAS"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            Select Case Rec1!VIA_CODIGO
                Case 1 ' comida
                    precio_comida = Chk0(Rec1!VIA_PRECIO)
                Case 2 ' Sopa
                    precio_sopa = Chk0(Rec1!VIA_PRECIO)
                Case 3 ' Postre
                    precio_postre = Chk0(Rec1!VIA_PRECIO)
                Case 4 ' Pan
                    precio_pan = Chk0(Rec1!VIA_PRECIO)
                Case 5 ' Descartable
                    precio_descartable = Chk0(Rec1!VIA_PRECIO)
                Case 6 ' Remise pilar
                    precio_remise = Chk0(Rec1!VIA_PRECIO)
                Case 7 ' Remise rio 2
                    precio_remise = Chk0(Rec1!VIA_PRECIO)
            End Select
            Rec1.MoveNext
        Loop
     End If
     Rec1.Close
End Function


Private Function cargar_comidas()
    sql = "SELECT COM_CODIGO,COM_DESCRI, TCOM_CODIGO "
    sql = sql & " FROM COMIDAS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
       Do While rec.EOF = False
          cboMenu.AddItem Trim(rec!COM_DESCRI)
          cboMenu.ItemData(cboMenu.NewIndex) = rec!COM_CODIGO
          
          
          cboVariantes.AddItem Trim(rec!COM_DESCRI)
          cboVariantes.ItemData(cboVariantes.NewIndex) = rec!COM_CODIGO
          rec.MoveNext
          
       Loop
       cboMenu.ListIndex = 0
       cboVariantes.ListIndex = 1
    End If
    rec.Close
End Function
Private Function preparargrillas()
    grdGrilla.FormatString = "Precio|<Fact|Nombre.|Almuer|Cena|>Sopa|Postre|Pan|Descartable|Remise|Vianda/Variantes|Diagnosticos|Entrega|Observaciones|Codigo Cliente|Monto|PrecioRemis"
    grdGrilla.ColWidth(0) = 700 'Precio
    grdGrilla.ColWidth(1) = 500 'Fact
    grdGrilla.ColWidth(2) = 2300 'Nombre
    grdGrilla.ColWidth(3) = 600 'Almuerzo
    grdGrilla.ColWidth(4) = 600  'Cena
    grdGrilla.ColWidth(5) = 600  'Sopa
    grdGrilla.ColWidth(6) = 600  'Postre
    grdGrilla.ColWidth(7) = 600  'Pan
    grdGrilla.ColWidth(8) = 600 'Descart.
    grdGrilla.ColWidth(9) = 600  'Remise
    grdGrilla.ColWidth(10) = 1800  'Vianda
    grdGrilla.ColWidth(11) = 1500  'Diagnostico
    grdGrilla.ColWidth(12) = 850  'Entrega
    grdGrilla.ColWidth(13) = 1000  'Observaciones
    grdGrilla.ColWidth(14) = 0  'Codigo cliente
    grdGrilla.ColWidth(15) = 0  'Monto inicial
    grdGrilla.ColWidth(16) = 0  'precio remis
    grdGrilla.Rows = 1
    grdGrilla.Cols = 17
    'grdGrilla.HighLight = flexHighlightNever
    grdGrilla.BorderStyle = flexBorderNone
    grdGrilla.row = 0
    For i = 0 To grdGrilla.Cols - 1
        grdGrilla.Col = i
        grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdGrilla.CellBackColor = &H808080    'GRIS OSCURO
        grdGrilla.CellFontBold = True
    Next
End Function
Private Function buscoplanilladiaria()

End Function

Private Sub grdGrilla_Click()
    Select Case grdGrilla.Col
    Case 3
        grdGrilla.ToolTipText = "Precio del almuerzo: $" & precio_comida
    Case 4
        grdGrilla.ToolTipText = "Precio de la cena: $" & precio_comida
    Case 5
        grdGrilla.ToolTipText = "Precio de la sopa: $" & precio_sopa
    Case 6
        grdGrilla.ToolTipText = "Precio de la postre: $" & precio_postre
    Case 7
        grdGrilla.ToolTipText = "Precio de la pan: $" & precio_pan
    Case 8
        grdGrilla.ToolTipText = "Precio de la descartable: $" & precio_descartable
    Case 9
        grdGrilla.ToolTipText = "Precio del remise: $" & precio_remise
    End Select
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1
            LimpiarFilasDeGrilla grdGrilla, grdGrilla.RowSel
            grdGrilla.Col = 0
            'txtSubtotal.Text = Valido_Importe(CStr(SumaTotal))
            'txtTotal.Text = Valido_Importe(CStr(SumaTotal))
            'txtPorcentajeIva_LostFocus
        Case 3, 4, 5, 6, 7, 8, 9, 10
            grdGrilla.TextMatrix(grdGrilla.RowSel, grdGrilla.Col) = ""
        End Select
        
    End If
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
            Case 3
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = "" Then
                    cmdGrabar.SetFocus
                End If
        End Select
    End If
End Sub

Private Sub grdGrilla_LeaveCell()
    If txtEdit.Visible = False Then Exit Sub
    grdGrilla = txtEdit.Text
    txtEdit.Visible = False
End Sub

Private Sub grdGrilla_LostFocus()
    If grdGrilla.Col = 9 Then
        grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = "$ " & grdGrilla.TextMatrix(grdGrilla.RowSel, 9)
    End If
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
       txtEdit.Visible = False
       grdGrilla.SetFocus
    End If
    If KeyCode = vbKeyDelete Then
        txtEdit.Visible = True
        txtEdit.Text = ""
        grdGrilla.SetFocus
    
    End If
    If KeyCode = vbKeyReturn Then
        'If grdGrilla.Col > 2 And grdGrilla.Col < 10 Then
        '    grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = monto(grdGrilla.RowSel)

        'End If
        mFoco = True
        grdGrilla.Col = 2
        grdGrilla.row = grdGrilla.RowSel
        grdGrilla.SetFocus
        
        
    End If
    grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = monto(grdGrilla.RowSel)
    'txtEdit.Visible = True
End Sub

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    If grdGrilla.Col = 0 Then KeyAscii = CarTexto(KeyAscii)
    If grdGrilla.Col = 3 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 4 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 5 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 6 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 7 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 8 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 9 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 10 Then KeyAscii = CarTexto(KeyAscii)
    
    
    CarTexto KeyAscii
End Sub
Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or (grdGrilla.Col = 2) Or (grdGrilla.Col = 3) Or (grdGrilla.Col = 4) Or (grdGrilla.Col = 5) Or (grdGrilla.Col = 6) Or (grdGrilla.Col = 7) Or (grdGrilla.Col = 8) Or (grdGrilla.Col = 9) Or (grdGrilla.Col = 10) Then
        If KeyAscii = vbKeyReturn Then
            If (grdGrilla.Col = 0) Or (grdGrilla.Col = 3) Or (grdGrilla.Col = 4) Or (grdGrilla.Col = 5) Or (grdGrilla.Col = 6) Or (grdGrilla.Col = 8) Or (grdGrilla.Col = 9) Or (grdGrilla.Col = 10) Then   '2
                If grdGrilla.row < grdGrilla.Rows - 1 Then
                    grdGrilla.row = grdGrilla.row + 1
                    grdGrilla.Col = 0
                Else
                    MySendKeys Chr(9)
                End If
            Else
                grdGrilla.Col = grdGrilla.Col + 1
            End If
        Else
            If (grdGrilla.Col = 3) Or (grdGrilla.Col = 4) Or (grdGrilla.Col = 5) Or (grdGrilla.Col = 6) Or (grdGrilla.Col = 7) Or (grdGrilla.Col = 8) Or (grdGrilla.Col = 9) Then  'grdGrilla.Col = 0 Or
                If KeyAscii > 47 And KeyAscii < 58 Then
                    EDITAR grdGrilla, txtEdit, KeyAscii
                End If
            ElseIf grdGrilla.Col = 1 Or grdGrilla.Col = 0 Or (grdGrilla.Col = 10) Then
                EDITAR grdGrilla, txtEdit, KeyAscii
            End If
        End If
    End If
End Sub
Private Function SumaTotal() As Double
    Dim i As Integer
    SumaTotal = 0
    For i = 1 To grdGrilla.Rows - 1
        SumaTotal = SumaTotal + CDbl(grdGrilla.TextMatrix(i, 0))
    Next
End Function
Private Function SumaTotalRemis() As Double
    Dim i As Integer
    Dim total As Double
    total = 0
    For i = 1 To grdGrilla.Rows - 1
        'If Chk0(grdGrilla.TextMatrix(i, 9)) = 1 Then  'tiene remis
            total = total + CDbl(Chk0(grdGrilla.TextMatrix(i, 16)))
        'End If
    Next
    SumaTotalRemis = total
End Function

