VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlanillaDiaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla Diaria"
   ClientHeight    =   8325
   ClientLeft      =   300
   ClientTop       =   1365
   ClientWidth     =   16035
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
   ScaleHeight     =   8325
   ScaleWidth      =   16035
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   11280
      TabIndex        =   1
      Top             =   7770
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   14895
      TabIndex        =   3
      Top             =   7770
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Cancelar"
      Height          =   450
      Left            =   13995
      TabIndex        =   2
      Top             =   7770
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7680
      Left            =   15
      TabIndex        =   11
      Top             =   30
      Width           =   15990
      _ExtentX        =   28205
      _ExtentY        =   13547
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameFactura"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmPlanillaDiaria.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameBuscar"
      Tab(1).Control(1).Enabled=   0   'False
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
         Height          =   6165
         Left            =   -74895
         TabIndex        =   21
         Top             =   1335
         Width           =   15795
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   300
            TabIndex        =   24
            Top             =   525
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   2775
            MaxLength       =   60
            TabIndex        =   23
            Top             =   5790
            Visible         =   0   'False
            Width           =   10290
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
            Left            =   13995
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   5700
            Width           =   1710
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   5490
            Left            =   75
            TabIndex        =   25
            Top             =   165
            Width           =   15630
            _ExtentX        =   27570
            _ExtentY        =   9684
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   1530
            TabIndex        =   27
            Top             =   5835
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   13155
            TabIndex        =   26
            Top             =   5760
            Width           =   645
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
         Left            =   195
         TabIndex        =   14
         Top             =   420
         Width           =   9990
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
            Format          =   109510657
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
            Left            =   7935
            MaskColor       =   &H000000FF&
            TabIndex        =   9
            ToolTipText     =   "Buscar "
            Top             =   915
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
            Format          =   109510657
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
            TabIndex        =   20
            Top             =   375
            Width           =   555
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Factura:"
            Height          =   195
            Left            =   1395
            TabIndex        =   19
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
         Left            =   -74895
         TabIndex        =   13
         Top             =   345
         Width           =   15720
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   12600
            TabIndex        =   34
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   109510657
            CurrentDate     =   43169
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar"
            Height          =   735
            Left            =   14160
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cboVariantes 
            Height          =   315
            ItemData        =   "frmPlanillaDiaria.frx":0038
            Left            =   1650
            List            =   "frmPlanillaDiaria.frx":003A
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   600
            Width           =   4000
         End
         Begin VB.ComboBox cboMenu 
            Height          =   315
            ItemData        =   "frmPlanillaDiaria.frx":003C
            Left            =   1650
            List            =   "frmPlanillaDiaria.frx":003E
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   240
            Width           =   4000
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
            Left            =   7200
            TabIndex        =   32
            Top             =   240
            Width           =   4455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Variantes:"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   31
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Menu:"
            Height          =   195
            Index           =   13
            Left            =   480
            TabIndex        =   29
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   12000
            TabIndex        =   18
            Top             =   300
            Width           =   495
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   5505
         Left            =   165
         TabIndex        =   10
         Top             =   1935
         Width           =   13860
         _ExtentX        =   24448
         _ExtentY        =   9710
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
      Left            =   13095
      TabIndex        =   0
      Top             =   7770
      Width           =   870
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
      Height          =   240
      Left            =   195
      TabIndex        =   17
      Top             =   7335
      Width           =   660
   End
End
Attribute VB_Name = "frmPlanillaDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label14_Click()

End Sub

Private Sub CmdBuscar_Click()
    crearplanilla
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmPlanillaDiaria = Nothing
        Unload Me
    End If
End Sub

Private Sub Fecha_Change()
    lblfecha.Caption = FechaLetras(Fecha.day, Fecha.dayofweek, Fecha.month, Fecha.year)
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
    sql = "SELECT  * FROM CLIENTE C, PROGRAMA_CLIENTE PC"
    sql = sql & " WHERE C.CLI_CODIGO=PC.CLI_CODIGO "
    sql = sql & " AND (PC.PRG_CODIGO = " & DIA & " OR PC.PRG_CODIGO= " & DIA + 8 & ")" 'ALMUERZO MAS CENA DE ESE
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Cli = 0
        i = 0
        Do While Not rec.EOF
            If rec!CLI_CODIGO = Cli Then
                'ACTUALIZAR COLUMMNA CENA
                grdGrilla.TextMatrix(i, 5) = rec!PRC_CANT
            Else
                Cli = rec!CLI_CODIGO
                grdGrilla.AddItem "$70,00" & Chr(9) & "M" & Chr(9) & rec!CLI_RAZSOC & Chr(9) & "" & Chr(9) & rec!PRC_CANT
                i = i + 1
            End If
            rec.MoveNext
        Loop
    End If
    rec.Close
    
'    sql = "SELECT  * FROM CLIENTE C, CLIENTE_VIANDAS CV,PROGRAMA_CLIENTE PG"
'    sql = sql & " WHERE C.CLI_CODIGO=CV.CLI_CODIGO "
'    sql = sql & " AND PG.CLI_CODIGO=C.CLI_CODIGO "
'    sql = sql & " AND (PC.PRG_CODIGO = " & DIA & " OR PC.PRG_CODIGO= " & DIA + 8 & ")" 'ALMUERZO MAS CENA DE ESE
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'
    
    'Recorro la grilla y cargo lo q se incluye de la vianda
    If grdGrilla.Rows > 1 Then
        For i = 1 To grdGrilla.Rows - 1
            
        Next
    End If
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
    
    
End Sub
Private Function cargar_comidas()
    sql = "SELECT COM_CODIGO,COM_DESCRI, TCOM_CODIGO "
    sql = sql & " FROM COMIDAS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
       Do While rec.EOF = False
          cboMenu.AddItem Trim(rec!COM_DESCRI)
          cboMenu.ItemData(cboMenu.NewIndex) = rec!COM_CODIGO
          rec.MoveNext
          
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
    grdGrilla.FormatString = "^Precio|<Fact|^Nombre.|>Sopa|Almuerza|Cena|>Vianda completa/Variantes|Postre|Pan|Diagnosticos|Entrega|Observaciones|Codigo Cliente"
    grdGrilla.ColWidth(0) = 700 'Precio
    grdGrilla.ColWidth(1) = 500 'Fact
    grdGrilla.ColWidth(2) = 2500 'Nombre
    grdGrilla.ColWidth(3) = 700 'Sopa
    grdGrilla.ColWidth(4) = 700  'Almuerzo
    grdGrilla.ColWidth(5) = 700  'Cena
    grdGrilla.ColWidth(6) = 3000 'Vianda
    grdGrilla.ColWidth(7) = 500  'Postre
    grdGrilla.ColWidth(8) = 500  'Pan
    grdGrilla.ColWidth(9) = 2000  'Diagnostico
    grdGrilla.ColWidth(10) = 700  'Entrega
    grdGrilla.ColWidth(11) = 3000  'Observaciones
    grdGrilla.ColWidth(12) = 0  'Codigo cliente
    grdGrilla.Rows = 1
    grdGrilla.Cols = 13
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

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1
            LimpiarFilasDeGrilla grdGrilla, grdGrilla.RowSel
            grdGrilla.Col = 0
            'txtSubtotal.Text = Valido_Importe(CStr(SumaTotal))
            'txtTotal.Text = Valido_Importe(CStr(SumaTotal))
            'txtPorcentajeIva_LostFocus
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

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mFoco = True
        grdGrilla.Col = 0
        grdGrilla.row = grdGrilla.RowSel
    End If
    If KeyCode = vbKeyEscape Then
       txtEdit.Visible = False
       grdGrilla.SetFocus
    End If
End Sub

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    If grdGrilla.Col = 0 Then KeyAscii = CarTexto(KeyAscii)
    If grdGrilla.Col = 3 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 4 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 5 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 7 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 8 Then KeyAscii = CarNumeroEntero(KeyAscii)
    
    
    CarTexto KeyAscii
End Sub
Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or (grdGrilla.Col = 2) Or (grdGrilla.Col = 3) Or (grdGrilla.Col = 4) Or (grdGrilla.Col = 5) Or (grdGrilla.Col = 7) Or (grdGrilla.Col = 8) Then
        If KeyAscii = vbKeyReturn Then
            If (grdGrilla.Col = 4) Or (grdGrilla.Col = 5) Or (grdGrilla.Col = 6) Or (grdGrilla.Col = 8) Or (grdGrilla.Col = 9) Then '2
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
            If (grdGrilla.Col = 3) Or (grdGrilla.Col = 4) Or (grdGrilla.Col = 5) Or (grdGrilla.Col = 7) Or (grdGrilla.Col = 8) Then 'grdGrilla.Col = 0 Or
                If KeyAscii > 47 And KeyAscii < 58 Then
                    EDITAR grdGrilla, txtEdit, KeyAscii
                End If
            ElseIf grdGrilla.Col = 1 Or grdGrilla.Col = 0 Then
                EDITAR grdGrilla, txtEdit, KeyAscii
            End If
        End If
    End If
End Sub
