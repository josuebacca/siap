VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ABMClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizacion de Clientes..."
   ClientHeight    =   8265
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleMode       =   0  'User
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabClientes 
      Height          =   7515
      Left            =   45
      TabIndex        =   11
      Top             =   45
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   13256
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "ABMClientes.frx":0BC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(12)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(11)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(10)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(9)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCuit"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtIngresosBrutos"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtMail"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtObserva"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboPais"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboIva"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "&Programas"
      TabPicture(1)   =   "ABMClientes.frx":0BDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Comidas"
      TabPicture(2)   =   "ABMClientes.frx":0BFA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame10"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Vianda"
      TabPicture(3)   =   "ABMClientes.frx":0C16
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtMedica"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).Control(2)=   "Frame4"
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame8 
         Caption         =   "Altas/Bajas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2655
         Left            =   120
         TabIndex        =   133
         Top             =   4680
         Width           =   7095
         Begin VB.Frame Frame11 
            Caption         =   "Evento"
            Height          =   2295
            Left            =   120
            TabIndex        =   136
            Top             =   240
            Width           =   3015
            Begin VB.TextBox txtmotivo 
               Height          =   795
               Left            =   240
               MaxLength       =   100
               MultiLine       =   -1  'True
               TabIndex        =   140
               ToolTipText     =   "Motivo de Alta o Baja"
               Top             =   1320
               Width           =   2685
            End
            Begin VB.OptionButton optAlta 
               Caption         =   "Alta"
               Height          =   255
               Left            =   240
               TabIndex        =   138
               Top             =   240
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton optBaja 
               Caption         =   "Baja"
               Height          =   255
               Left            =   1560
               TabIndex        =   137
               Top             =   240
               Width           =   735
            End
            Begin MSComCtl2.DTPicker fechaevento 
               Height          =   315
               Left            =   240
               TabIndex        =   139
               Top             =   720
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   102825985
               CurrentDate     =   40071
            End
         End
         Begin VB.CommandButton cmdAgregaEvento 
            Caption         =   ">"
            Height          =   375
            Left            =   3240
            TabIndex        =   134
            Top             =   1200
            Width           =   375
         End
         Begin MSFlexGridLib.MSFlexGrid grillaBaja 
            Height          =   2445
            Left            =   3720
            TabIndex        =   135
            Top             =   120
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   4313
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorSel    =   16761024
            AllowBigSelection=   -1  'True
            FocusRect       =   0
            HighLight       =   2
            SelectionMode   =   1
            AllowUserResizing=   3
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
      End
      Begin VB.Frame Frame6 
         Caption         =   "Diagnostico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1335
         Left            =   -74880
         TabIndex        =   131
         Top             =   480
         Width           =   7095
         Begin VB.TextBox txtDiagnostico 
            Height          =   915
            Left            =   120
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   132
            Top             =   240
            Width           =   6885
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo Facturacion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4815
         Left            =   -70800
         TabIndex        =   67
         Top             =   480
         Width           =   2775
         Begin VB.OptionButton optSemanal 
            Caption         =   "Semanal"
            Height          =   255
            Left            =   960
            TabIndex        =   73
            Top             =   1320
            Width           =   1215
         End
         Begin VB.OptionButton optQuincenal 
            Caption         =   "Quincenal"
            Height          =   255
            Left            =   960
            TabIndex        =   72
            Top             =   1920
            Width           =   1215
         End
         Begin VB.OptionButton optMensual 
            Caption         =   "Mensual"
            Height          =   255
            Left            =   960
            TabIndex        =   71
            Top             =   2520
            Width           =   1215
         End
         Begin VB.OptionButton optDiaria 
            Caption         =   "Diaria"
            Height          =   255
            Left            =   960
            TabIndex        =   70
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Viandas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4815
         Left            =   -74880
         TabIndex        =   61
         Top             =   480
         Width           =   3975
         Begin VB.Frame Frame1 
            Caption         =   "Frame7"
            Height          =   3975
            Left            =   240
            TabIndex        =   62
            Top             =   480
            Width           =   3495
            Begin VB.TextBox txtImpVianda 
               BackColor       =   &H8000000B&
               Height          =   315
               Index           =   6
               Left            =   1920
               TabIndex        =   119
               Text            =   "0,00"
               Top             =   2880
               Width           =   735
            End
            Begin VB.TextBox txtImpVianda 
               BackColor       =   &H8000000B&
               Height          =   315
               Index           =   5
               Left            =   1920
               TabIndex        =   118
               Text            =   "0,00"
               Top             =   2520
               Width           =   735
            End
            Begin VB.TextBox txtImpVianda 
               BackColor       =   &H8000000B&
               Height          =   315
               Index           =   4
               Left            =   1920
               TabIndex        =   117
               Text            =   "0,00"
               Top             =   2160
               Width           =   735
            End
            Begin VB.TextBox txtImpVianda 
               BackColor       =   &H8000000B&
               Height          =   315
               Index           =   3
               Left            =   1920
               TabIndex        =   116
               Text            =   "0,00"
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox txtImpVianda 
               BackColor       =   &H8000000B&
               Height          =   315
               Index           =   2
               Left            =   1920
               TabIndex        =   115
               Text            =   "0,00"
               Top             =   1440
               Width           =   735
            End
            Begin VB.TextBox txtImpVianda 
               BackColor       =   &H8000000B&
               Height          =   315
               Index           =   1
               Left            =   1920
               TabIndex        =   114
               Text            =   "0,00"
               Top             =   1080
               Width           =   735
            End
            Begin VB.CheckBox chkVianda 
               Alignment       =   1  'Right Justify
               Caption         =   "Otros"
               Height          =   255
               Index           =   6
               Left            =   480
               TabIndex        =   113
               Top             =   2880
               Width           =   1185
            End
            Begin VB.CheckBox chkVianda 
               Alignment       =   1  'Right Justify
               Caption         =   "Remise"
               Height          =   255
               Index           =   5
               Left            =   480
               TabIndex        =   112
               Top             =   2520
               Width           =   1185
            End
            Begin VB.CheckBox chkVianda 
               Alignment       =   1  'Right Justify
               Caption         =   "Pan"
               Height          =   255
               Index           =   4
               Left            =   480
               TabIndex        =   111
               Top             =   2160
               Width           =   1185
            End
            Begin VB.CheckBox chkVianda 
               Alignment       =   1  'Right Justify
               Caption         =   "Postre"
               Height          =   255
               Index           =   3
               Left            =   480
               TabIndex        =   110
               Top             =   1800
               Width           =   1185
            End
            Begin VB.CheckBox chkVianda 
               Alignment       =   1  'Right Justify
               Caption         =   "Sopa"
               Height          =   255
               Index           =   2
               Left            =   480
               TabIndex        =   109
               Top             =   1440
               Width           =   1185
            End
            Begin VB.CheckBox chkVianda 
               Alignment       =   1  'Right Justify
               Caption         =   "Descartable"
               Height          =   255
               Index           =   1
               Left            =   480
               TabIndex        =   108
               Top             =   1080
               Width           =   1185
            End
            Begin VB.TextBox txtTotalVianda 
               BackColor       =   &H00FFC0C0&
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
               Left            =   1920
               TabIndex        =   68
               Top             =   3360
               Width           =   735
            End
            Begin VB.CheckBox chkVianda 
               Alignment       =   1  'Right Justify
               Caption         =   "Comida"
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   65
               Top             =   720
               Width           =   1185
            End
            Begin VB.CheckBox chktodas 
               Caption         =   "Marcar Todos"
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   0
               Width           =   1575
            End
            Begin VB.TextBox txtImpVianda 
               BackColor       =   &H8000000B&
               Height          =   315
               Index           =   0
               Left            =   1920
               TabIndex        =   63
               Text            =   "0,00"
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "TOTAL:"
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
               Left            =   1200
               TabIndex        =   69
               Top             =   3420
               Width           =   585
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Importe"
               Height          =   195
               Left            =   1920
               TabIndex        =   66
               Top             =   360
               Width           =   570
            End
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Programa por Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   6855
         Left            =   -74880
         TabIndex        =   48
         Top             =   480
         Width           =   7095
         Begin VB.Frame Frame7 
            Height          =   4575
            Left            =   120
            TabIndex        =   51
            Top             =   2160
            Width           =   3255
            Begin VB.TextBox txtAlm 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   7
               Left            =   1320
               MaxLength       =   4
               TabIndex        =   105
               Top             =   4080
               Width           =   495
            End
            Begin VB.CheckBox chkAlm 
               Alignment       =   1  'Right Justify
               Caption         =   "Feriados"
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   104
               Top             =   4080
               Width           =   945
            End
            Begin VB.TextBox txtCena 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   7
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   103
               Top             =   4080
               Width           =   495
            End
            Begin VB.TextBox txtCena 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   6
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   102
               Top             =   3600
               Width           =   495
            End
            Begin VB.TextBox txtCena 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   5
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   101
               Top             =   3120
               Width           =   495
            End
            Begin VB.TextBox txtCena 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   100
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtCena 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   99
               Top             =   2160
               Width           =   495
            End
            Begin VB.TextBox txtCena 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   98
               Top             =   1680
               Width           =   495
            End
            Begin VB.TextBox txtCena 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   97
               Top             =   1200
               Width           =   495
            End
            Begin VB.CheckBox chkCena 
               Caption         =   "Check1"
               Height          =   255
               Index           =   7
               Left            =   2040
               TabIndex        =   96
               Top             =   4080
               Width           =   255
            End
            Begin VB.CheckBox chkCena 
               Caption         =   "Check1"
               Height          =   255
               Index           =   6
               Left            =   2040
               TabIndex        =   95
               Top             =   3600
               Width           =   255
            End
            Begin VB.CheckBox chkCena 
               Caption         =   "Check1"
               Height          =   255
               Index           =   5
               Left            =   2040
               TabIndex        =   94
               Top             =   3120
               Width           =   255
            End
            Begin VB.CheckBox chkCena 
               Caption         =   "Check1"
               Height          =   255
               Index           =   4
               Left            =   2040
               TabIndex        =   93
               Top             =   2640
               Width           =   255
            End
            Begin VB.CheckBox chkCena 
               Caption         =   "Check1"
               Height          =   255
               Index           =   3
               Left            =   2040
               TabIndex        =   92
               Top             =   2160
               Width           =   255
            End
            Begin VB.CheckBox chkCena 
               Caption         =   "Check1"
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   91
               Top             =   1680
               Width           =   255
            End
            Begin VB.CheckBox chkCena 
               Caption         =   "Check1"
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   90
               Top             =   1200
               Width           =   255
            End
            Begin VB.TextBox txtCena 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   2280
               MaxLength       =   4
               TabIndex        =   89
               Top             =   720
               Width           =   495
            End
            Begin VB.CheckBox chkCena 
               Caption         =   "Check1"
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   88
               Top             =   720
               Width           =   255
            End
            Begin VB.TextBox txtAlm 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   6
               Left            =   1320
               MaxLength       =   4
               TabIndex        =   87
               Top             =   3600
               Width           =   495
            End
            Begin VB.CheckBox chkAlm 
               Alignment       =   1  'Right Justify
               Caption         =   "Domingo"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   86
               Top             =   3600
               Width           =   945
            End
            Begin VB.TextBox txtAlm 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   5
               Left            =   1320
               MaxLength       =   4
               TabIndex        =   85
               Top             =   3120
               Width           =   495
            End
            Begin VB.CheckBox chkAlm 
               Alignment       =   1  'Right Justify
               Caption         =   "Sábado"
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   84
               Top             =   3120
               Width           =   945
            End
            Begin VB.TextBox txtAlm 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               Left            =   1320
               MaxLength       =   4
               TabIndex        =   83
               Top             =   2640
               Width           =   495
            End
            Begin VB.TextBox txtAlm 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               Left            =   1320
               MaxLength       =   4
               TabIndex        =   82
               Top             =   2160
               Width           =   495
            End
            Begin VB.TextBox txtAlm 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   1320
               MaxLength       =   4
               TabIndex        =   81
               Top             =   1680
               Width           =   495
            End
            Begin VB.TextBox txtAlm 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   1320
               MaxLength       =   4
               TabIndex        =   80
               Top             =   1200
               Width           =   495
            End
            Begin VB.CheckBox chkAlm 
               Alignment       =   1  'Right Justify
               Caption         =   "Viernes"
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   79
               Top             =   2640
               Width           =   945
            End
            Begin VB.CheckBox chkAlm 
               Alignment       =   1  'Right Justify
               Caption         =   "Jueves"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   78
               Top             =   2160
               Width           =   945
            End
            Begin VB.CheckBox chkAlm 
               Alignment       =   1  'Right Justify
               Caption         =   "Miercoles"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   77
               Top             =   1680
               Width           =   1065
            End
            Begin VB.TextBox txtAlm 
               BackColor       =   &H8000000B&
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   1320
               MaxLength       =   4
               TabIndex        =   76
               Top             =   720
               Width           =   495
            End
            Begin VB.CheckBox chkAlm 
               Alignment       =   1  'Right Justify
               Caption         =   "Martes"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   75
               Top             =   1200
               Width           =   945
            End
            Begin VB.CheckBox chkAlm 
               Alignment       =   1  'Right Justify
               Caption         =   "Lunes"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   74
               Top             =   720
               Width           =   945
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Almuerzo"
               Height          =   195
               Left            =   960
               TabIndex        =   55
               Top             =   360
               Width           =   660
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Cena"
               Height          =   195
               Left            =   2040
               TabIndex        =   52
               Top             =   360
               Width           =   375
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Dias Cancelados"
            Height          =   4575
            Left            =   3480
            TabIndex        =   49
            Top             =   2160
            Width           =   3495
            Begin VB.TextBox txtCom 
               Height          =   315
               Left            =   120
               TabIndex        =   54
               Top             =   720
               Width           =   3015
            End
            Begin VB.CommandButton cmdAgregaCan 
               Caption         =   "Agregar"
               Height          =   375
               Left            =   2400
               TabIndex        =   56
               Top             =   240
               Width           =   735
            End
            Begin MSFlexGridLib.MSFlexGrid grillaCan 
               Height          =   3285
               Left            =   120
               TabIndex        =   50
               Top             =   1200
               Width           =   3120
               _ExtentX        =   5503
               _ExtentY        =   5794
               _Version        =   393216
               Cols            =   6
               FixedCols       =   0
               RowHeightMin    =   300
               BackColorSel    =   16761024
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               HighLight       =   2
               SelectionMode   =   1
               AllowUserResizing=   3
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
            Begin MSComCtl2.DTPicker FecCanc 
               Height          =   315
               Left            =   120
               TabIndex        =   53
               Top             =   360
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   102825985
               CurrentDate     =   40070
            End
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Comidas por Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   5415
         Left            =   -74880
         TabIndex        =   43
         Top             =   1920
         Width           =   7095
         Begin VB.ComboBox cboTipoComida 
            Height          =   315
            ItemData        =   "ABMClientes.frx":0C32
            Left            =   2400
            List            =   "ABMClientes.frx":0C34
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   360
            Width           =   2445
         End
         Begin VB.CommandButton cmdAgTodos 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3300
            TabIndex        =   47
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmdAgUno 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3300
            TabIndex        =   46
            ToolTipText     =   "Agregar Comida Seleccionada"
            Top             =   1960
            Width           =   495
         End
         Begin VB.CommandButton cmdQuUno 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3300
            TabIndex        =   45
            ToolTipText     =   "Quitar Comida Seleccionada"
            Top             =   2480
            Width           =   495
         End
         Begin VB.CommandButton cmdQuTodos 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3300
            TabIndex        =   44
            Top             =   3000
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSFlexGridLib.MSFlexGrid grillaComidas 
            Height          =   4245
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   7488
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorSel    =   16761024
            AllowBigSelection=   -1  'True
            FocusRect       =   0
            HighLight       =   2
            SelectionMode   =   1
            AllowUserResizing=   3
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
         Begin MSFlexGridLib.MSFlexGrid grillaComCli 
            Height          =   4245
            Left            =   3840
            TabIndex        =   58
            Top             =   1080
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   7488
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorSel    =   16761024
            AllowBigSelection=   -1  'True
            FocusRect       =   0
            HighLight       =   2
            SelectionMode   =   1
            AllowUserResizing=   3
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Grupos:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1800
            TabIndex        =   107
            Top             =   360
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Comidas seleccionadas por el cliente"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3840
            TabIndex        =   60
            Top             =   840
            Width           =   2595
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Comidas disponibles"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   840
            Width           =   1425
         End
      End
      Begin VB.TextBox txtMedica 
         Height          =   285
         Left            =   -64440
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   6380
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox cboIva 
         Height          =   315
         ItemData        =   "ABMClientes.frx":0C36
         Left            =   7200
         List            =   "ABMClientes.frx":0C38
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   5420
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.ComboBox cboPais 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "ABMClientes.frx":0C3A
         Left            =   7200
         List            =   "ABMClientes.frx":0C3C
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   5060
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4215
         Left            =   120
         TabIndex        =   20
         Top             =   390
         Width           =   7095
         Begin VB.TextBox txtNombre 
            Height          =   315
            Left            =   1290
            MaxLength       =   50
            TabIndex        =   129
            Top             =   630
            Width           =   4000
         End
         Begin VB.TextBox txtOcupacion 
            Height          =   315
            Left            =   1290
            MaxLength       =   100
            TabIndex        =   128
            Top             =   3015
            Width           =   4000
         End
         Begin VB.TextBox txtEdad 
            Height          =   315
            Left            =   4170
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   127
            Top             =   2610
            Width           =   1155
         End
         Begin VB.ComboBox cboProvincia 
            Height          =   315
            ItemData        =   "ABMClientes.frx":0C3E
            Left            =   1290
            List            =   "ABMClientes.frx":0C40
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   1800
            Width           =   4000
         End
         Begin VB.ComboBox cboLocalidad 
            Height          =   315
            ItemData        =   "ABMClientes.frx":0C42
            Left            =   1290
            List            =   "ABMClientes.frx":0C44
            Style           =   2  'Dropdown List
            TabIndex        =   125
            Top             =   2190
            Width           =   4000
         End
         Begin VB.TextBox txtTelefono 
            Height          =   315
            Left            =   1290
            MaxLength       =   30
            TabIndex        =   124
            Top             =   3405
            Width           =   4000
         End
         Begin VB.TextBox txtFax 
            Height          =   315
            Left            =   1290
            MaxLength       =   30
            TabIndex        =   123
            Top             =   3795
            Width           =   4000
         End
         Begin VB.TextBox txtDomicilio 
            Height          =   315
            Left            =   1290
            MaxLength       =   50
            TabIndex        =   122
            Top             =   1020
            Width           =   4000
         End
         Begin VB.TextBox txtCodPostal 
            Height          =   315
            Left            =   1290
            MaxLength       =   10
            TabIndex        =   121
            Top             =   1410
            Width           =   1155
         End
         Begin VB.TextBox txtID 
            Height          =   315
            Left            =   1290
            TabIndex        =   120
            Top             =   240
            Width           =   1155
         End
         Begin VB.TextBox txtNroDoc 
            Height          =   315
            Left            =   4140
            MaxLength       =   9
            TabIndex        =   1
            Top             =   240
            Width           =   1155
         End
         Begin VB.TextBox txtimagen 
            Height          =   405
            Left            =   5520
            TabIndex        =   42
            Top             =   3360
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.CommandButton cmdFotos 
            Caption         =   "Cargar Foto"
            Height          =   375
            Left            =   5400
            TabIndex        =   6
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtBuscaOS 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2610
            MaxLength       =   40
            TabIndex        =   3
            Top             =   6285
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txtNAfiliado 
            Height          =   315
            Left            =   2610
            MaxLength       =   25
            TabIndex        =   5
            Top             =   6660
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txtBuscarOSNombre 
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
            Left            =   3705
            MaxLength       =   50
            TabIndex        =   4
            Tag             =   "Descripción"
            Top             =   6285
            Visible         =   0   'False
            Width           =   2900
         End
         Begin VB.CommandButton cmdBuscaOS 
            Height          =   315
            Left            =   7080
            Picture         =   "ABMClientes.frx":0C46
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Buscar Obras Sociales"
            Top             =   5325
            Visible         =   0   'False
            Width           =   400
         End
         Begin VB.TextBox txtDNI 
            Height          =   315
            Left            =   8940
            MaxLength       =   10
            TabIndex        =   0
            Top             =   4080
            Visible         =   0   'False
            Width           =   1275
         End
         Begin MSComCtl2.DTPicker DTFechaPCons 
            Height          =   315
            Left            =   2610
            TabIndex        =   2
            Top             =   6000
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102825985
            CurrentDate     =   40071
         End
         Begin MSComCtl2.DTPicker DTFechaNac 
            Height          =   315
            Left            =   1290
            TabIndex        =   130
            Top             =   2610
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   102825985
            CurrentDate     =   40071
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1575
            Left            =   5400
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   37
            Top             =   1860
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Doc.:"
            Height          =   195
            Index           =   3
            Left            =   3240
            TabIndex        =   36
            Top             =   270
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código Postal:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   1455
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "F. Nacimiento:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Top             =   2670
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   33
            Top             =   1065
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Celular:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   32
            Top             =   3870
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   31
            Top             =   3480
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Localidad:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   30
            Top             =   2250
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   660
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Id.:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   270
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Edad:"
            Height          =   195
            Index           =   14
            Left            =   3000
            OLEDropMode     =   1  'Manual
            TabIndex        =   27
            Top             =   2670
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "F Primer Cons.:"
            Height          =   195
            Index           =   0
            Left            =   1440
            TabIndex        =   26
            Top             =   6060
            Width           =   1110
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Obra Social:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   1440
            TabIndex        =   25
            Top             =   6315
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Afiliado :"
            Height          =   195
            Index           =   15
            Left            =   1440
            TabIndex        =   24
            Top             =   6705
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ocupacion:"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   23
            Top             =   3075
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "DNI:"
            Height          =   195
            Index           =   25
            Left            =   8520
            TabIndex        =   22
            Top             =   4140
            Visible         =   0   'False
            Width           =   330
         End
      End
      Begin VB.TextBox txtObserva 
         Height          =   810
         Left            =   9945
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   5815
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtMail 
         Height          =   315
         Left            =   9930
         MaxLength       =   50
         TabIndex        =   13
         Top             =   5350
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtIngresosBrutos 
         Height          =   315
         Left            =   9840
         MaxLength       =   10
         TabIndex        =   12
         Top             =   4120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSMask.MaskEdBox txtCuit 
         Height          =   315
         Left            =   9480
         TabIndex        =   14
         Top             =   4735
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   13
         Mask            =   "##-########-#"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cond. I.V.A.:"
         Height          =   195
         Index           =   9
         Left            =   5640
         TabIndex        =   40
         Top             =   6040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "e-mail:"
         Height          =   195
         Index           =   7
         Left            =   9240
         TabIndex        =   18
         Top             =   5395
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C.U.I.T.:"
         Height          =   195
         Index           =   10
         Left            =   9480
         TabIndex        =   17
         Top             =   5080
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ing. Brutos:"
         Height          =   195
         Index           =   11
         Left            =   9960
         TabIndex        =   16
         Top             =   4480
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observación:"
         Height          =   195
         Index           =   12
         Left            =   8880
         TabIndex        =   15
         Top             =   5815
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   2880
      TabIndex        =   10
      Top             =   7680
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   2520
      Picture         =   "ABMClientes.frx":1648
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   585
      Left            =   6240
      Picture         =   "ABMClientes.frx":1792
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7635
      Width           =   1065
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   585
      Left            =   5040
      Picture         =   "ABMClientes.frx":1A9C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7635
      Width           =   1065
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccionar Perfil"
      FileName        =   "Perfil1.jpg"
      Filter          =   "*.jgp, *.bmp, *.gif"
      InitDir         =   "...\"
      Orientation     =   2
   End
End
Attribute VB_Name = "ABMClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'parametros para la configuración de la ventana de datos
Dim vFieldID As String
Dim vStringSQL As String
Dim vFormLlama As Form
Dim vMode As Integer
Dim vListView As ListView
Dim vDesFieldID As String
Dim Pais As String
Dim Provincia As String
Dim i As Integer
Dim nCCRowSel As Integer
Dim nCMRowSel As Integer
Dim ActivoGrid As Integer ' 1 actio 0 desactivo

'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "CLIENTE"
Const cCampoID = "CLI_CODIGO"
Const cDesRegistro = "Cliente"

Function ActualizarListaBase(pMode As Integer)
    On Error GoTo moco
    Dim rec As ADODB.Recordset
    Dim cSQL As String
    Dim i As Integer
    Dim auxListItem As ListItem
    Dim IndiceCampoID As Integer
    Dim OrdenCampo As Integer
    Dim f As ADODB.Field
    Set rec = New ADODB.Recordset
    
    'armo la cadena a ejecutar
    If InStr(1, vStringSQL, "WHERE") = 0 Then
        cSQL = vStringSQL & " WHERE " & cCampoID & " = " & txtID.Text
    Else
        cSQL = vStringSQL & " AND " & cCampoID & " = " & txtID.Text
    End If
    
    If pMode = 4 Then
        vListView.ListItems.Remove vListView.SelectedItem.Index
        Exit Function
    End If
    
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
        If rec.EOF = False Then
        
'            'busco el indce del campo identificador
            OrdenCampo = 0
            IndiceCampoID = 0
            For Each f In rec.Fields
                OrdenCampo = OrdenCampo + 1
                If UCase(f.Name) = UCase(vDesFieldID) Then
                    IndiceCampoID = OrdenCampo - 1
                End If
            Next f
        
            'recorro la coleción de campos a actualizar
            For i = 0 To rec.Fields.Count - 1
                If i = 0 Then
                    Select Case pMode
                        Case 1
                            Set auxListItem = vListView.ListItems.Add(, "'" & rec.Fields(IndiceCampoID) & "'", CStr(IIf(IsNull(rec.Fields(i)), "", rec.Fields(i))), 1)
                            auxListItem.Icon = 1
                            auxListItem.SmallIcon = 1
                            
                        Case 2
                            Set auxListItem = vListView.SelectedItem
                            auxListItem.Text = rec.Fields(i)
                    End Select
                Else
                    auxListItem.SubItems(i) = IIf(IsNull(rec.Fields(i)), "", rec.Fields(i))
                End If
            Next i
        End If
    End If
    Exit Function
moco:
    If Err.Number = 35613 Then
        Call Menu.mnuContextABM_Click(4)
    End If
End Function

Function SetMode(pMode As Integer)

    'Configura los controles del form segun el parametro pMode
    'Parametro: pMode indica el modo en que se utilizará este form
    '  pMode  =             1> Indica nuevo registro
    '                       2> Editar registro existente
    '                       3> Mostrar dato del registro existente
    '                       4> Eliminar registro existente
    
    
    Select Case pMode
        Case 1, 2
            AcCtrl txtDNI
            AcCtrlx txtNombre
            AcCtrlx cboIva
            AcCtrlx txtCuit
            AcCtrlx txtIngresosBrutos
            AcCtrlx txtNroDoc
            AcCtrlx DTFechaNac
            'AcCtrlx cboPais
            
            AcCtrlx cboProvincia
            AcCtrlx cboLocalidad
            AcCtrlx txtDomicilio
            AcCtrlx txtTelefono
            AcCtrlx txtFax
            AcCtrlx txtCodPostal
            AcCtrlx txtMail
            AcCtrlx txtObserva
            
            AcCtrlx txtEdad
            txtEdad.Locked = True
            AcCtrlx txtOcupacion
            AcCtrlx DTFechaPCons
'            For i = 0 To 7
'                AcCtrlx chkAlm(i)
'                AcCtrlx chkCena(i)
'                AcCtrlx txtAlm(i)
'                AcCtrlx txtCena(i)
'            Next i
            AcCtrlx FecCanc
            AcCtrlx cmdAceptar
            AcCtrlx txtCom
            AcCtrlx grillaCan
            
'            AcCtrlx txtBuscaOS
'            AcCtrlx cmdBuscaOS
'            AcCtrlx txtBuscarOSNombre
'            AcCtrlx txtNAfiliado
'
'            AcCtrlx txtMC
'            AcCtrlx txtRelac
'            AcCtrlx txtAFA
'            AcCtrlx txtAPP
'            AcCtrlx txtEFisico
'            AcCtrlx txtDiag
'            AcCtrlx txtEstCom
'            AcCtrlx txtPTest
'            AcCtrlx txtHC
'
'            AcCtrlx txtMedica
'            AcCtrlx cmdFotos
'
'            'Anamnesis
'            AcCtrlx chkTomaMed
'            AcCtrlx txtCualMe
'            AcCtrlx txtAlergia
'            AcCtrlx chkAneste
'            AcCtrlx chktuhemo
'            AcCtrlx chktarcic
'            AcCtrlx chkDiabet
'            AcCtrlx chkprealt
'            AcCtrlx chkprebaj
'            AcCtrlx chkEpilep
'            AcCtrlx chkEmbara
'            AcCtrlx txtMeses
'            AcCtrlx chkLactan
'            AcCtrlx chkhemofi
'            AcCtrlx chkcardia
'            AcCtrlx txtcualca
'            AcCtrlx chkmarcapaso
'            AcCtrlx DTUltVis
'            AcCtrlx cboAnamTrat
'            AcCtrlx txtcuadia

            
        Case 3, 4
            DesacCtrl txtDNI
            DesacCtrlx txtNombre
            DesacCtrlx cboIva
            DesacCtrlx txtCuit
            DesacCtrlx txtIngresosBrutos
            DesacCtrlx txtNroDoc
            DesacCtrlx DTFechaNac
            'DesacCtrlx cboPais
            DesacCtrlx cboProvincia
            DesacCtrlx cboLocalidad
            DesacCtrlx txtDomicilio
            DesacCtrlx txtTelefono
            DesacCtrlx txtFax
            DesacCtrlx txtCodPostal
            DesacCtrlx txtMail
            DesacCtrlx txtObserva
            DesacCtrlx txtEdad
            DesacCtrlx txtOcupacion
'            For i = 0 To 7
'                DesacCtrlx chkAlm(i)
'                DesacCtrlx chkCena(i)
'                DesacCtrlx txtAlm(i)
'                DesacCtrlx txtCena(i)
'            Next i
            DesacCtrlx FecCanc
            DesacCtrlx cmdAceptar
            DesacCtrlx txtCom
            DesacCtrlx grillaCan
            
'            DesacCtrlx DTFechaPCons
'            DesacCtrlx txtBuscaOS
'            DesacCtrlx cmdBuscaOS
'            DesacCtrlx txtBuscarOSNombre
'            DesacCtrlx txtNAfiliado
'
'            DesacCtrlx txtMC
'            DesacCtrlx txtRelac
'            DesacCtrlx txtAFA
'            DesacCtrlx txtAPP
'            DesacCtrlx txtEFisico
'            DesacCtrlx txtDiag
'            DesacCtrlx txtEstCom
'            DesacCtrlx txtPTest
'            DesacCtrlx txtHC
'
'            DesacCtrlx txtMedica
'            DesacCtrlx cmdFotos
'
'            'Anamnesis
'            DesacCtrlx chkTomaMed
'            DesacCtrlx txtCualMe
'            DesacCtrlx txtAlergia
'            DesacCtrlx chkAneste
'            DesacCtrlx chktuhemo
'            DesacCtrlx chktarcic
'            DesacCtrlx chkDiabet
'            DesacCtrlx chkprealt
'            DesacCtrlx chkprebaj
'            DesacCtrlx chkEpilep
'            DesacCtrlx chkEmbara
'            DesacCtrlx txtMeses
'            DesacCtrlx chkLactan
'            DesacCtrlx chkhemofi
'            DesacCtrlx chkcardia
'            DesacCtrlx txtcualca
'            DesacCtrlx chkmarcapaso
'            DesacCtrlx DTUltVis
'            DesacCtrlx cboAnamTrat
'            DesacCtrlx txtcuadia
            
            
    End Select
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nuevo " & cDesRegistro
            txtID_LostFocus
            DesacCtrl txtID
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando " & cDesRegistro & " - " & Trim(txtNombre)
            DesacCtrl txtID
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos del " & cDesRegistro
            DesacCtrl txtID
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando " & cDesRegistro
            DesacCtrl txtID
    End Select
    
End Function

Public Function SetWindow(pWindow As Form, pSQL As String, pMode As Integer, pListview As ListView, pDesID As String)
    
    Set vFormLlama = pWindow 'Objeto ventana que que llama a la ventana de datos
    vStringSQL = pSQL 'string utilizado para argar la lista base
    vMode = pMode  'modo en que se utilizará la ventana de datos
    Set vListView = pListview 'objeto listview que se está editando
    vDesFieldID = pDesID 'nombre del campo identificador
    
    'valor del campo identificador de registro seleccionado (0 si es un reg. nuevo)
    If vMode <> 1 Then
        If vListView.SelectedItem.Selected = True Then
            vFieldID = vListView.SelectedItem.Key
        Else
            vFieldID = 0
        End If
    Else
        vFieldID = 0
    End If

End Function


Function Validar(pMode As Integer) As Boolean

    If pMode <> 4 Then
        Validar = False
        If txtID.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Identificación del  " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        ElseIf txtNombre.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Nombre del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtNombre.SetFocus
            Exit Function
        
        ElseIf cboPais.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Paí del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboPais.SetFocus
            Exit Function
            
        ElseIf cboProvincia.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Provincia del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboPais.SetFocus
            Exit Function
        
        ElseIf cboLocalidad.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Localidad del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboProvincia.SetFocus
            Exit Function
        End If
    End If
    
    Validar = True
    
End Function

Private Sub cboCanal_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboAnamTrat_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboAnamTrat_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboDoctor_Change()
    'cmdAceptar.Enabled = True
End Sub

Private Sub cboIva_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboLocalidad_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboPais_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboPais_LostFocus()
    If vMode = 2 And Pais = cboPais.Text Then
        Exit Sub
    End If
    Set Rec1 = New ADODB.Recordset
    cboProvincia.Clear
    sql = "SELECT PRO_CODIGO,PRO_DESCRI"
    sql = sql & " FROM PROVINCIA "
    sql = sql & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
    'sql = sql & " AND PRO_CODIGO=1" 'CORDOBA
    sql = sql & " ORDER BY PRO_DESCRI"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If (Rec1.BOF And Rec1.EOF) = 0 Then
       Do While Rec1.EOF = False
          cboProvincia.AddItem Trim(Rec1!PRO_DESCRI)
          cboProvincia.ItemData(cboProvincia.NewIndex) = Rec1!PRO_CODIGO
          Rec1.MoveNext
       Loop
       cboProvincia.ListIndex = cboProvincia.ListIndex + 1
       BuscaProx "CORDOBA", cboProvincia
    Else
       MsgBox "No hay cargado Provincia para ese País.", vbOKOnly + vbCritical, TIT_MSGBOX
    End If
    Rec1.Close
    cboProvincia_LostFocus
End Sub

Private Sub cboProvincia_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboProvincia_LostFocus()
    If vMode = 2 And Provincia = cboProvincia.Text Then
        Exit Sub
    End If
    Set Rec1 = New ADODB.Recordset
    cboLocalidad.Clear
    sql = "SELECT LOC_CODIGO,LOC_DESCRI FROM LOCALIDAD"
    sql = sql & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
    sql = sql & " AND PRO_CODIGO=" & cboProvincia.ItemData(cboProvincia.ListIndex)
    sql = sql & " ORDER BY LOC_DESCRI "
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If (Rec1.BOF And Rec1.EOF) = 0 Then
       Do While Rec1.EOF = False
          cboLocalidad.AddItem Trim(Rec1!LOC_DESCRI)
          cboLocalidad.ItemData(cboLocalidad.NewIndex) = Rec1!LOC_CODIGO
          Rec1.MoveNext
       Loop
       cboLocalidad.ListIndex = cboLocalidad.ListIndex + 1
       BuscaProx "PILAR", cboLocalidad
    Else
       MsgBox "No hay cargada Localidad para esta Provincia.", vbOKOnly + vbCritical, TIT_MSGBOX
    End If
    Rec1.Close
    'BuscaProx "CORDOBA", cboLocalidad
End Sub

Private Sub cboTratamiento_Change()
    'cmdAceptar.Enabled = True
End Sub

Private Sub cboTipoComida_LostFocus()
    If cboTipoComida.Text = "<TODOS>" Then
        CargarGrillaComida 0
    Else
        CargarGrillaComida cboTipoComida.ItemData(cboTipoComida.ListIndex)
    End If
End Sub

Private Sub chkAlm_Click(Index As Integer)
    cmdAceptar.Enabled = True
    If chkAlm(Index).Value = Checked Then
        txtAlm(Index).BackColor = &HFFFFFF
        txtAlm(Index) = 1
        txtAlm(Index).Enabled = True
    Else
        txtAlm(Index).BackColor = &H8000000B
        txtAlm(Index) = ""
        txtAlm(Index).Enabled = False
    End If
End Sub
Private Sub chkCena_Click(Index As Integer)
    cmdAceptar.Enabled = True
    If chkCena(Index).Value = Checked Then
        txtCena(Index).BackColor = &HFFFFFF
        txtCena(Index) = 1
        txtCena(Index).Enabled = True
    Else
        txtCena(Index).BackColor = &H8000000B
        txtCena(Index) = ""
        txtCena(Index).Enabled = False
    End If
End Sub

Private Sub chkVianda_Click(Index As Integer)
    cmdAceptar.Enabled = True
    If chkVianda(Index).Value = Checked Then
        txtImpVianda(Index).BackColor = &HFFFFFF
        txtImpVianda(Index) = "1,00" ' aca buscar el precio
        txtImpVianda(Index).Enabled = True
    Else
        txtImpVianda(Index).BackColor = &H8000000B
        txtImpVianda(Index) = "0,00"
        txtImpVianda(Index).Enabled = False
    End If
    sumarVianda
End Sub
Private Sub sumarVianda()
    Dim totVianda As Double
    totVianda = 0
    For i = 0 To 6
        totVianda = totVianda + txtImpVianda(i).Text
    Next i
    txtTotalVianda.Text = Format(totVianda, "#,##")
    
End Sub

Private Sub cmdAceptar_Click()

    Dim cSQL As String
    Dim cSQLAnam As String
    Dim cli_fac As Integer
    cli_fac = 0
    If optDiaria.Value = True Then cli_fac = 1
    If optSemanal.Value = True Then cli_fac = 2
    If optQuincenal.Value = True Then cli_fac = 3
    If optMensual.Value = True Then cli_fac = 4
    If Validar(vMode) = True Then
        
        On Error GoTo ErrorTran
        
        Screen.MousePointer = vbHourglass
    
        DBConn.BeginTrans
        Select Case vMode
            Case 1 'nuevo
            
                cSQL = "INSERT INTO " & cTabla
                cSQL = cSQL & "     (CLI_CODIGO, CLI_RAZSOC, CLI_DNI, CLI_DOMICI, CLI_CUIT,"
                cSQL = cSQL & " CLI_INGBRU, "
                If Not IsNull(DTFechaNac.Value) Then
                    cSQL = cSQL & " CLI_CUMPLE, "
                End If
                cSQL = cSQL & " IVA_CODIGO, CLI_NRODOC,"
                cSQL = cSQL & " CLI_TELEFONO, CLI_MAIL, CLI_FAX, CLI_CODPOS,"
                cSQL = cSQL & " LOC_CODIGO, PRO_CODIGO, PAI_CODIGO, CLI_OBSERVA, "
                cSQL = cSQL & " CLI_EDAD, CLI_OCUPACION, "
                
                If Not IsNull(DTFechaPCons.Value) Then
                    cSQL = cSQL & "CLI_FECPC,"
                End If
                
                cSQL = cSQL & "OS_NUMERO,CLI_NROAFIL,CLI_FACTURA,CLI_DIAGNO) "
                
'                cSQL = cSQL & " CLI_MC, CLI_RELAC, CLI_AFA,CLI_APP,CLI_EFISICO, "
'                cSQL = cSQL & " CLI_DIAG, CLI_ESTCOM, CLI_PTEST,CLI_HC,CLI_MEDICA,CLI_FOTO,CLI_ASPCLI,CLI_FACTURA) "
                
                cSQL = cSQL & " VALUES "
                cSQL = cSQL & "     (" & XN(txtID.Text) & ", " & XS(txtNombre.Text) & ", "
                cSQL = cSQL & XN(txtDNI.Text) & ", "
                cSQL = cSQL & XS(txtDomicilio.Text) & ", " & XS(txtCuit.Text) & ", "
                cSQL = cSQL & XS(txtIngresosBrutos.Text) & ", "
                
                If Not IsNull(DTFechaNac.Value) Then
                    cSQL = cSQL & XDQ(DTFechaNac.Value) & ", "
                End If
                
                cSQL = cSQL & cboIva.ItemData(cboIva.ListIndex) & ", "
                cSQL = cSQL & XN(txtNroDoc.Text) & ", "
                cSQL = cSQL & XS(txtTelefono.Text) & ", "
                cSQL = cSQL & XS(txtMail.Text) & ", " & XS(txtFax.Text) & ", "
                cSQL = cSQL & XS(txtCodPostal.Text) & ", "
                cSQL = cSQL & cboLocalidad.ItemData(cboLocalidad.ListIndex) & ", "
                cSQL = cSQL & cboProvincia.ItemData(cboProvincia.ListIndex) & ", "
                cSQL = cSQL & cboPais.ItemData(cboPais.ListIndex) & ","
                cSQL = cSQL & XS(Trim(txtObserva.Text)) & ","
                cSQL = cSQL & XN(txtEdad.Text) & ", "
                cSQL = cSQL & XS(Trim(txtOcupacion.Text)) & ","
                
                If Not IsNull(DTFechaPCons.Value) Then
                    cSQL = cSQL & XDQ(DTFechaPCons.Value) & ", "
                End If
                
                cSQL = cSQL & XN(txtBuscaOS.Text) & ", "
                cSQL = cSQL & XS(txtNAfiliado.Text) & ", "
                
                'AGREGAR el tipo de facturacion del cliente

                
                cSQL = cSQL & cli_fac & ", "
                cSQL = cSQL & XS(txtDiagnostico) & ") "
                

                'sql = Insertar PROGRAMA POR CLIENTE
                
                For i = 0 To 7
                    If chkAlm(i).Value = Checked Then
                        sql = "INSERT INTO PROGRAMA_CLIENTE (CLI_CODIGO,PRG_CODIGO,PRC_CANT)"
                        sql = sql & " VALUES("
                        sql = sql & XN(txtID) & ","
                        If i = 6 Then ' es domingo
                            sql = sql & 1 & ","
                        Else
                            If i = 7 Then
                                sql = sql & i + 1 & ","
                            Else
                                sql = sql & i + 2 & ","
                            End If
                        End If
                        sql = sql & XN(txtAlm(i)) & ")"
                        DBConn.Execute sql
                    End If
                    If chkCena(i).Value = Checked Then
                        sql = "INSERT INTO PROGRAMA_CLIENTE (CLI_CODIGO,PRG_CODIGO,PRC_CANT)"
                        sql = sql & " VALUES("
                        sql = sql & XN(txtID) & ","
                        If i = 6 Then ' es domingo
                            sql = sql & 9 & ","
                        Else
                            If i = 7 Then
                                sql = sql & i + 9 & ","
                            Else
                                sql = sql & i + 10 & ","
                            End If
                        End If
                        sql = sql & XN(txtCena(i)) & ")"
                        DBConn.Execute sql
                    End If
                Next i
                 
                'Agregar dias CANCELADOS
                For i = 1 To grillaCan.Rows - 1
                    sql = "INSERT INTO CANCELA_CLIENTE (CLI_CODIGO,CAC_FECHA,CAC_COMEN)"
                    sql = sql & " VALUES("
                    sql = sql & XN(txtID) & ","
                    sql = sql & XDQ(grillaCan.TextMatrix(i, 0)) & ","
                    sql = sql & XS(grillaCan.TextMatrix(i, 1)) & ")"
                    DBConn.Execute sql
                Next i
                
                ' Agregar COMIDAS POR CLIENTE
                
                For i = 1 To grillaComCli.Rows - 1
                    sql = "INSERT INTO COMIDAS_CLIENTE (CLI_CODIGO,COM_CODIGO)"
                    sql = sql & " VALUES("
                    sql = sql & XN(txtID) & ","
                    sql = sql & XN(grillaComCli.TextMatrix(i, 2)) & ")"
                    DBConn.Execute sql
                Next i
                
                ' AGREGAR VIANDA
                For i = 0 To 6
                    If txtImpVianda(i).Text > 0 Then
                        sql = "INSERT INTO CLIENTE_VIANDAS (CLI_CODIGO,VIA_CODIGO)"
                        sql = sql & " VALUES("
                        sql = sql & XN(txtID) & ","
                        sql = sql & XN(i + 1) & ")"
                        DBConn.Execute sql
                    End If
                Next i
                
                'registrar altas/bajas
'                For i = 1 To grillaComCli.Rows - 1
'                    sql = "INSERT INTO CLIENTE_ALTABAJA (CAB_CODIGO,CLI_CODIGO,CAB_FECHA,CAB_EVENTO,CAB_MOTIVO)"
'                    sql = sql & " VALUES("
'                    sql = sql & XN(i) & ","
'                    sql = sql & XN(txtID) & ","
'                    sql = sql & XDQ(grillaBaja.TextMatrix(i, 0)) & ","
'                    sql = sql & XS(grillaBaja.TextMatrix(i, 1)) & ","
'                    sql = sql & XS(grillaBaja.TextMatrix(i, 2)) & ")"
'                    DBConn.Execute sql
'                Next i
                
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "  CLI_RAZSOC=" & XS(txtNombre.Text)
                cSQL = cSQL & " ,CLI_DNI=" & XS(txtDNI.Text)
                cSQL = cSQL & " ,CLI_DOMICI=" & XS(txtDomicilio.Text)
                cSQL = cSQL & " ,CLI_CUIT=" & XS(txtCuit.Text)
                cSQL = cSQL & " ,CLI_INGBRU=" & XS(txtIngresosBrutos.Text)
                If Not IsNull(DTFechaNac.Value) Then
                    cSQL = cSQL & " ,CLI_CUMPLE=" & XDQ(DTFechaNac.Value)
                End If
                cSQL = cSQL & " ,IVA_CODIGO=" & cboIva.ItemData(cboIva.ListIndex)
                cSQL = cSQL & " ,CLI_TELEFONO=" & XS(txtTelefono.Text)
                cSQL = cSQL & " ,CLI_MAIL=" & XS(txtMail.Text)
                cSQL = cSQL & " ,CLI_FAX=" & XS(txtFax.Text)
                cSQL = cSQL & " ,CLI_CODPOS=" & XS(txtCodPostal.Text)
                cSQL = cSQL & " ,LOC_CODIGO=" & cboLocalidad.ItemData(cboLocalidad.ListIndex)
                cSQL = cSQL & " ,PRO_CODIGO=" & cboProvincia.ItemData(cboProvincia.ListIndex)
                cSQL = cSQL & " ,PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
                cSQL = cSQL & " ,CLI_OBSERVA=" & XS(Trim(txtObserva.Text))
                cSQL = cSQL & " ,CLI_NRODOC=" & XN(txtNroDoc.Text)
                cSQL = cSQL & " ,CLI_EDAD= " & XN(txtEdad.Text)
                cSQL = cSQL & " ,CLI_OCUPACION=" & XS(Trim(txtOcupacion.Text))
                If Not IsNull(DTFechaPCons.Value) Then
                    cSQL = cSQL & " ,CLI_FECPC=" & XDQ(DTFechaPCons.Value)
                End If
                'OS_NUMERO,CLI_NROAFIL
                cSQL = cSQL & " ,OS_NUMERO= " & XN(txtBuscaOS.Text)
                cSQL = cSQL & " ,CLI_NROAFIL=" & XS(txtNAfiliado.Text)
                cSQL = cSQL & " ,CLI_FACTURA=" & cli_fac
                cSQL = cSQL & " ,CLI_DIAGNO=" & XS(txtDiagnostico.Text)
                cSQL = cSQL & " WHERE CLI_CODIGO  = " & XN(txtID.Text)
                DBConn.Execute cSQL
                
               'Modificar PROGRAMA POR CLIENTE
                sql = "DELETE FROM PROGRAMA_CLIENTE WHERE CLI_CODIGO = " & XN(txtID)
                DBConn.Execute sql
                
                For i = 0 To 7
                    If chkAlm(i).Value = Checked Then
                        sql = "INSERT INTO PROGRAMA_CLIENTE (CLI_CODIGO,PRG_CODIGO,PRC_CANT)"
                        sql = sql & " VALUES("
                        sql = sql & XN(txtID) & ","
                        If i = 6 Then ' es domingo
                            sql = sql & 1 & ","
                        Else
                            If i = 7 Then
                                sql = sql & i + 1 & ","
                            Else
                                sql = sql & i + 2 & ","
                            End If
                        End If
                        sql = sql & XN(txtAlm(i)) & ")"
                        DBConn.Execute sql
                    End If
                    If chkCena(i).Value = Checked Then
                        sql = "INSERT INTO PROGRAMA_CLIENTE (CLI_CODIGO,PRG_CODIGO,PRC_CANT)"
                        sql = sql & " VALUES("
                        sql = sql & XN(txtID) & ","
                        If i = 6 Then ' es domingo
                            sql = sql & 9 & ","
                        Else
                            If i = 7 Then
                                sql = sql & i + 9 & ","
                            Else
                                sql = sql & i + 10 & ","
                            End If
                        End If
                        sql = sql & XN(txtCena(i)) & ")"
                        DBConn.Execute sql
                    End If
                Next i
                
                ' Modificar dias CANCELADOS
                sql = "DELETE FROM CANCELA_CLIENTE WHERE CLI_CODIGO = " & XN(txtID)
                DBConn.Execute sql
                For i = 1 To grillaCan.Rows - 1
                    sql = "INSERT INTO CANCELA_CLIENTE (CLI_CODIGO,CAC_FECHA,CAC_COMEN)"
                    sql = sql & " VALUES("
                    sql = sql & XN(txtID) & ","
                    sql = sql & XDQ(grillaCan.TextMatrix(i, 0)) & ","
                    sql = sql & XS(grillaCan.TextMatrix(i, 1)) & ")"
                    DBConn.Execute sql
                Next i
                
                ' Modificar COMIDAS POR CLIENTE
                sql = "DELETE FROM COMIDAS_CLIENTE WHERE CLI_CODIGO = " & XN(txtID)
                DBConn.Execute sql
                For i = 1 To grillaComCli.Rows - 1
                    sql = "INSERT INTO COMIDAS_CLIENTE (CLI_CODIGO,COM_CODIGO)"
                    sql = sql & " VALUES("
                    sql = sql & XN(txtID) & ","
                    sql = sql & XN(grillaComCli.TextMatrix(i, 2)) & ")"
                    DBConn.Execute sql
                Next i
                
                ' MODIFICAR VIANDA
                                
                sql = "DELETE FROM CLIENTE_VIANDAS WHERE CLI_CODIGO = " & XN(txtID)
                DBConn.Execute sql
                For i = 0 To 6
                    If txtImpVianda(i).Text > 0 Then
                        sql = "INSERT INTO CLIENTE_VIANDAS (CLI_CODIGO,VIA_CODIGO)"
                        sql = sql & " VALUES("
                        sql = sql & XN(txtID) & ","
                        sql = sql & XN(i + 1) & ")"
                        DBConn.Execute sql
                    End If
                Next i
                
                
                'Modificar altas/bajas
                'DBConn.Execute "DELETE FROM CLIENTE_ALTABAJA WHERE CLI_CODIGO = " & XN(txtID)
                                
'                For i = 1 To grillaComCli.Rows - 1
'                    sql = "INSERT INTO CLIENTE_ALTABAJA (CAB_CODIGO,CLI_CODIGO,CAB_FECHA,CAB_EVENTO,CAB_MOTIVO)"
'                    sql = sql & " VALUES("
'                    sql = sql & XN(i) & ","
'                    sql = sql & XN(txtID) & ","
'                    sql = sql & XDQ(grillaBaja.TextMatrix(i, 0)) & ","
'                    sql = sql & XS(grillaBaja.TextMatrix(i, 1)) & ","
'                    sql = sql & XS(grillaBaja.TextMatrix(i, 2)) & ")"
'                    DBConn.Execute sql
'                Next i
                
            Case 4 'eliminar
                cSQL = "DELETE FROM " & cTabla & " WHERE CLI_CODIGO  = " & XN(txtID.Text)
                
                'cSQLAnam = "DELETE FROM CLIENTE_ANAM WHERE CLI_CODIGO  = " & XN(txtID.Text)
                
        End Select
        
        DBConn.Execute cSQL

        
        'DBConn.Execute sql
        DBConn.CommitTrans
        'On Error GoTo 0
        
        'actualizo la lista base
        ActualizarListaBase vMode
        
        Screen.MousePointer = vbDefault
        Unload Me
    End If
    Exit Sub
    
ErrorTran:
    
    DBConn.RollbackTrans
    Screen.MousePointer = vbDefault
    
    'manejo el error
    'ManejoDeErrores DBConn.ErrorNative
    MsgBox Err.Description, vbCritical
    
End Sub
'Private Function InsertAnamnesis() As String
'    sql = "INSERT INTO CLIENTE_ANAM"
'    sql = sql & " (CLI_CODIGO, CLA_TOMMED,CLA_CUALME,CLA_ALERGIA,CLA_ANESTE, "
'    sql = sql & "CLA_TUHEMO,CLA_TARCIC,CLA_DIABET,CLA_PREALT,CLA_PREBAJ,"
'    sql = sql & "CLA_EPILEP,CLA_EMBARA,CLA_MESES,"
'    sql = sql & "CLA_LACTAN,CLA_HEMOFI,CLA_CARDIA,CLA_CUALCA, "
'    sql = sql & "CLA_MARCAP,"
'    If Not IsNull(DTUltVis.Value) Then
'        sql = sql & "CLA_ULTVIS, "
'    End If
'    sql = sql & "TR_CODIGO,CLA_CUADIA,CLA_OTROS)"
'
'    sql = sql & " VALUES "
'    sql = sql & "     (" & XN(txtID.Text) & ", " & chkTomaMed.Value & ", "
'    sql = sql & XS(txtCualMe.Text) & ", "
'    sql = sql & XS(txtAlergia.Text) & ", " & chkAneste.Value & ", "
'    sql = sql & chktuhemo.Value & ", "
'    sql = sql & chktarcic.Value & ", "
'    sql = sql & chkDiabet.Value & ", "
'    sql = sql & chkprealt.Value & ", "
'    sql = sql & chkprebaj.Value & ", "
'    sql = sql & chkEpilep.Value & ", "
'    sql = sql & chkEmbara.Value & ", " & XN(txtMeses.Text) & ", "
'    sql = sql & chkLactan.Value & ", "
'    sql = sql & chkhemofi.Value & ", "
'    sql = sql & chkcardia.Value & ", "
'    sql = sql & XS(Trim(txtcualca.Text)) & ","
'    sql = sql & chkmarcapaso.Value & ", "
'    If Not IsNull(DTUltVis.Value) Then
'        sql = sql & XDQ(DTUltVis.Value) & ","
'    End If
'    If cboAnamTrat.ListIndex <> 0 Then
'        sql = sql & cboAnamTrat.ItemData(cboAnamTrat.ListIndex) & ", "
'    Else
'        sql = sql & "0" & ", "
'    End If
'    sql = sql & XN(txtcuadia.Text) & ", "
'    sql = sql & XS(txtAnamOtros.Text) & ") "
'
'    InsertAnamnesis = sql
'
'End Function
'Private Function ActualizarAnamnesis() As String
'    sql = "UPDATE CLIENTE_ANAM SET "
'    sql = sql & "CLA_TOMMED = " & chkTomaMed.Value
'    sql = sql & ",CLA_CUALME = " & XS(txtCualMe.Text)
'    sql = sql & ",CLA_ALERGIA = " & XS(txtAlergia.Text)
'    sql = sql & ",CLA_ANESTE = " & chkAneste.Value
'    sql = sql & ",CLA_TUHEMO = " & chktuhemo.Value
'    sql = sql & ",CLA_TARCIC = " & chktarcic.Value
'    sql = sql & ",CLA_DIABET = " & chkDiabet.Value
'    sql = sql & ",CLA_PREALT = " & chkprealt.Value
'    sql = sql & ",CLA_PREBAJ = " & chkprebaj.Value
'    sql = sql & ",CLA_EPILEP = " & chkEpilep.Value
'    sql = sql & ",CLA_EMBARA = " & chkEmbara.Value
'    sql = sql & ",CLA_MESES = " & XN(txtMeses.Text)
'    sql = sql & ",CLA_LACTAN = " & chkLactan.Value
'    sql = sql & ",CLA_HEMOFI =" & chkhemofi.Value
'    sql = sql & ",CLA_CARDIA = " & chkcardia.Value
'    sql = sql & ",CLA_CUALCA = " & XS(Trim(txtcualca.Text))
'    sql = sql & ",CLA_MARCAP = " & chkmarcapaso.Value
'    If Not IsNull(DTUltVis.Value) Then
'        sql = sql & ",CLA_ULTVIS = " & XDQ(DTUltVis.Value)
'    End If
'    sql = sql & ",TR_CODIGO = " & cboAnamTrat.ItemData(cboAnamTrat.ListIndex)
'    sql = sql & ",CLA_CUADIA = " & XN(txtcuadia.Text)
'    sql = sql & ",CLA_OTROS = " & XS(txtAnamOtros.Text)
'    sql = sql & " WHERE CLI_CODIGO = " & XN(txtID.Text)
'
'    ActualizarAnamnesis = sql
'
'End Function
'Private Function validarcclinico() As Boolean
'    If DTFecha.Value = "" Then
'        MsgBox "Debe ingresar la Fecha", vbExclamation, TIT_MSGBOX
'        DTFecha.SetFocus
'        validarcclinico = False
'        Exit Function
'    End If
'    If cboDoctor.ListIndex = -1 Then
'        MsgBox "Debe seleccionar el Doctor", vbExclamation, TIT_MSGBOX
'        cboDoctor.SetFocus
'        validarcclinico = False
'        Exit Function
'    End If
'    'If txtCodTra.Text = "" Then
'    '    MsgBox "Debe ingresar el Tratamiento", vbExclamation, TIT_MSGBOX
'    '    txtCodTra.SetFocus
'    '    validarcclinico = False
'    '    Exit Function
'    'End If
'    If txtIndicaciones.Text = "" Then
'        MsgBox "Debe ingresar las Observaciones", vbExclamation, TIT_MSGBOX
'        txtIndicaciones.SetFocus
'        validarcclinico = False
'        Exit Function
'    End If
'    validarcclinico = True
'
'End Function
'
'Private Function validarcmedica() As Boolean
'    If DTMedFec.Value = "" Then
'        MsgBox "Debe ingresar la Fecha", vbExclamation, TIT_MSGBOX
'        DTMedFec.SetFocus
'        validarcmedica = False
'        Exit Function
'    End If
'    If cboMedDoc.ListIndex = -1 Then
'        MsgBox "Debe seleccionar el Doctor", vbExclamation, TIT_MSGBOX
'        cboMedDoc.SetFocus
'        validarcmedica = False
'        Exit Function
'    End If
'    If cboMedica.ListIndex = -1 Then
'        MsgBox "Debe seleccionar el Medicamento", vbExclamation, TIT_MSGBOX
'        cboMedica.SetFocus
'        validarcmedica = False
'        Exit Function
'    End If
'    If txtMedIndica.Text = "" Then
'        MsgBox "Debe ingresar las Indicaciones", vbExclamation, TIT_MSGBOX
'        txtMedIndica.SetFocus
'        validarcmedica = False
'        Exit Function
'    End If
'    validarcmedica = True
'
'End Function
'Private Function CargarCClinico(paciente As Integer)
'
'    sql = "SELECT CC.*, D.VEN_NOMBRE,T.TR_DESCRI,T.TR_CODNUE "
'    sql = sql & " FROM CCLINICO CC, VENDEDOR D,TRATAMIENTO T"
'    sql = sql & " WHERE D.VEN_CODIGO = CC.VEN_CODIGO"
'    sql = sql & " AND T.TR_CODIGO = CC.TR_CODIGO"
'    sql = sql & " AND CC.CLI_CODIGO = " & paciente
'    sql = sql & " ORDER BY CC.CCL_FECHA DESC"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    grdCClinico.Rows = 1
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'            grdCClinico.AddItem rec!CCL_FECHA & Chr(9) & rec!TR_DESCRI & Chr(9) & _
'                            rec!CCL_INDICA & Chr(9) & rec!VEN_NOMBRE & Chr(9) & _
'                            rec!TR_CODIGO & Chr(9) & _
'                            rec!VEN_CODIGO & Chr(9) & rec!CCL_NUMERO & Chr(9) & _
'                            rec!CCL_FECPC & Chr(9) & rec!TR_CODNUE
'            rec.MoveNext
'        Loop
'    End If
'    rec.Close
'End Function
'Private Function CargarCMedica(paciente As Integer)
'
'    sql = "SELECT CC.*, D.VEN_NOMBRE,T.MED_NOMBRE "
'    sql = sql & " FROM CMEDICA CC, VENDEDOR D,MEDICAMENTOS T"
'    sql = sql & " WHERE D.VEN_CODIGO = CC.VEN_CODIGO"
'    sql = sql & " AND T.MED_CODIGO = CC.MED_CODIGO"
'    sql = sql & " AND CC.CLI_CODIGO = " & paciente
'    sql = sql & " ORDER BY CC.CME_FECHA DESC"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    GrdCMedica.Rows = 1
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'            GrdCMedica.AddItem rec!CME_FECHA & Chr(9) & rec!MED_NOMBRE & Chr(9) & _
'                            rec!CME_INDICA & Chr(9) & rec!VEN_NOMBRE & Chr(9) & _
'                            rec!MED_CODIGO & Chr(9) & _
'                            rec!VEN_CODIGO & Chr(9) & rec!CME_NUMERO
'            rec.MoveNext
'        Loop
'    End If
'    rec.Close
'End Function
'Private Sub cmdAgregar_Click()
'    Dim nMaxCCodigo As Integer
'    If validarcclinico = False Then Exit Sub
'    If MsgBox("¿Confirma el Curso Clinico?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'
'    On Error GoTo HayErrorCClinico
'        DBConn.BeginTrans
'
'
'
'        If txtCCodigo.Text = "" Then
'
'        ' Nuevo Curso Clinico
'            rec.Open "SELECT MAX(CCL_NUMERO) AS MAXIMO FROM CCLINICO", DBConn, adOpenStatic, adLockOptimistic
'            If rec.EOF = False Then
'                nMaxCCodigo = IIf(IsNull(rec!Maximo), 1, rec!Maximo + 1)
'            End If
'            rec.Close
'            sql = "INSERT INTO CCLINICO"
'            sql = sql & " (CCL_NUMERO,CCL_FECHA, CLI_CODIGO,VEN_CODIGO,"
'            sql = sql & " TR_CODIGO,"
'
'            If Not IsNull(DTFecPC.Value) Then
'                sql = sql & " CCL_FECPC,"
'            End If
'
'            sql = sql & " CCL_INDICA)"
'            sql = sql & " VALUES ("
'            sql = sql & nMaxCCodigo & ","
'            sql = sql & XDQ(DTFecha.Value) & ","
'            sql = sql & txtID & ","
'            sql = sql & cboDoctor.ItemData(cboDoctor.ListIndex) & ","
'            sql = sql & XN(txtIdTra) & ","
'            If Not IsNull(DTFecPC.Value) Then
'                sql = sql & XDQ(DTFecPC.Value) & ","
'            End If
'            sql = sql & XS(txtIndicaciones.Text) & ")"
'
'        Else
'        ' Modifico Curso Clinico
'            If MsgBox("¿Confirma la Modificación del Curso Clinico?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'            ' aca hago el update
'            sql = "UPDATE CCLINICO SET "
'            sql = sql & " CCL_FECHA = " & XDQ(DTFecha.Value)
'            sql = sql & " ,CLI_CODIGO =" & XN(txtID)
'            sql = sql & " ,VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
'            sql = sql & " ,TR_CODIGO = " & XN(txtIdTra)
'            If Not IsNull(DTFecPC.Value) Then
'                sql = sql & ",CCL_FECPC = " & XDQ(DTFecPC.Value)
'            End If
'            sql = sql & " ,CCL_INDICA =" & XS(txtIndicaciones.Text)
'            sql = sql & " WHERE CCL_NUMERO = " & XN(txtCCodigo.Text)
'
'        End If
'        DBConn.Execute sql
'        DBConn.CommitTrans
'        CargarCClinico txtID.Text
'        LimpiarCClinico
'
'    Exit Sub
'
'HayErrorCClinico:
'    Screen.MousePointer = vbNormal
'    If rec.State = 1 Then rec.Close
'    If Rec1.State = 1 Then Rec1.Close
'    DBConn.RollbackTrans
'    MsgBox Err.Description, vbCritical, TIT_MSGBOX
'
'End Sub
'Private Function LimpiarCClinico()
'    nCCRowSel = 0
'    DTFecha.Value = Date
'
'    If User <> 99 Then
'        Call BuscaCodigoProxItemData(XN(User), cboDoctor)
'    Else
'        cboDoctor.ListIndex = -1
'    End If
'
'    'cboTratamiento.ListIndex = -1
'    txtIdTra.Text = ""
'    txtCodTra.Text = ""
'    txtDescTra.Text = ""
'    txtIndicaciones.Text = ""
'    txtAspCli.Text = ""
'    txtCCodigo.Text = ""
'    DTFecPC.Value = Null
'End Function
'Private Function LimpiarCMedica()
'    nCMRowSel = 0
'    DTMedFec.Value = Date
'    If User <> 99 Then
'        Call BuscaCodigoProxItemData(XN(User), cboMedDoc)
'    Else
'        cboDoctor.ListIndex = -1
'    End If
'    cboMedica.ListIndex = -1
'    txtMedIndica.Text = ""
'    txtMedCodigo.Text = ""
'End Function
'Private Sub cmdAyuda_Click()
'    Call WinHelp(Me.hWnd, App.Path & "\help\AYUDA.HLP", cdlHelpContext, 12)
'End Sub
'
'Private Sub cmdBuscaOS_Click()
''    frmBuscar.TipoBusqueda = 8
''    frmBuscar.TxtDescriB = ""
''    frmBuscar.Show vbModal
''    If frmBuscar.grdBuscar.Text <> "" Then
''        frmBuscar.grdBuscar.Col = 0
''        txtBuscaOS.Text = frmBuscar.grdBuscar.Text
''        frmBuscar.grdBuscar.Col = 1
''        txtBuscarOSNombre.Text = frmBuscar.grdBuscar.Text
''    Else
''        txtBuscaOS.SetFocus
''    End If
'
'End Sub

Private Sub cmdAgregaCan_Click()
    If FecCanc.Value <> "" Then
        'VERIFICAR Q NO ESTE ESA FECHA
        For i = 1 To grillaCan.Rows - 1
            If grillaCan.TextMatrix(i, 0) = Format(FecCanc.Value, "DD/MM/YYYY") Then
                MsgBox "El día ya fue seleccionado como CANCELADO", vbInformation, TIT_MSGBOX
                Exit Sub
            End If
        Next i
        grillaCan.AddItem Format(FecCanc.Value, "DD/MM/YYYY") & Chr(9) & txtCom
    End If
End Sub

Private Sub cmdAgregaEvento_Click()
    Dim evento As String
    If optAlta.Value = True Then
        evento = "Alta"
    Else
        evento = "Baja"
    End If
    If fechaevento.Value = "" Then
        MsgBox "Ingrese la fecha de " & evento & ".", vbExclamation, TIT_MSGBOX
        fechaevento.SetFocus
        Exit Sub
    End If
    If txtmotivo.Text = "" Then
        MsgBox "Ingrese el motivo de " & evento & ".", vbExclamation, TIT_MSGBOX
        txtmotivo.SetFocus
        Exit Sub
    End If
        'VERIFICAR Q NO ESTE ESA FECHA
'        For i = 1 To grillaBaja.Rows - 1
'            If grillaBaja.TextMatrix(i, 0) = Format(fechaevento.Value, "DD/MM/YYYY") Then
'                MsgBox "El día ya fue seleccionado como CANCELADO", vbInformation, TIT_MSGBOX
'                Exit Sub
'            End If
'        Next i
    If grillaBaja.Rows = 1 And evento = "Baja" Then
        MsgBox "Debe ingresar el Alta antes que una baja", vbExclamation, TIT_MSGBOX
        optAlta.SetFocus
        Exit Sub
    End If
    grillaBaja.AddItem Format(fechaevento.Value, "DD/MM/YYYY") & Chr(9) & evento & Chr(9) & txtmotivo.Text
        
    'End If
End Sub

Private Sub cmdAgTodos_Click()
    For i = 1 To grillaCan.Rows - 1
        If grillaComCli.TextMatrix(i, 2) = grillaComidas.TextMatrix(i, 2) Then
            MsgBox "La comida ya fue seleccionada para este Cliente ", vbInformation, TIT_MSGBOX
            Exit Sub
        End If
    Next i
End Sub

Private Sub cmdAgUno_Click()
    cmdAceptar.Enabled = True
    'validar que exista
    For i = 1 To grillaComCli.Rows - 1
        If grillaComCli.TextMatrix(i, 2) = grillaComidas.TextMatrix(grillaComidas.RowSel, 2) Then
            MsgBox "La comida ya fue seleccionado para este Cliente ", vbInformation, TIT_MSGBOX
            Exit Sub
        End If
    Next i
    
    grillaComCli.AddItem grillaComidas.TextMatrix(grillaComidas.RowSel, 0) & Chr(9) & _
                         grillaComidas.TextMatrix(grillaComidas.RowSel, 1) & Chr(9) & _
                         grillaComidas.TextMatrix(grillaComidas.RowSel, 2) & Chr(9) & _
                         grillaComidas.TextMatrix(grillaComidas.RowSel, 3)


End Sub

Private Sub cmdCerrar_Click()

    Unload Me
    
End Sub

Private Sub cmdFotos_Click()
    cmdAceptar.Enabled = True
    On Error Resume Next
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Seleccione un nombre de archivo"
    CommonDialog1.Filter = "Pictures(*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
    
    CommonDialog1.ShowOpen
    If Err.Number = 0 Then
        If CommonDialog1.FileName Like "*.bmp" _
        Or CommonDialog1.FileName Like "*.gif" _
        Or CommonDialog1.FileName Like "*.jpg" Then
            
            Image1.Picture = LoadPicture(CommonDialog1.FileName)
            txtimagen.Text = CommonDialog1.FileName
            On Error GoTo 0
        Else
            MsgBox "El Archivo seleccionado no es válido", vbExclamation, Me.Caption
        End If
    End If
End Sub

'Private Sub cmdMedAgregar_Click()
'    Dim nMaxCCodigo As Integer
'    If validarcmedica = False Then Exit Sub
'    If MsgBox("¿Confirma la Medicacion?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'
'    On Error GoTo HayErrorCMedica
'        DBConn.BeginTrans
'
'
'        If txtMedCodigo.Text = "" Then
'
'        ' Nueva Medicacion
'            rec.Open "SELECT MAX(CME_NUMERO) AS MAXIMO FROM CMEDICA", DBConn, adOpenStatic, adLockOptimistic
'            If rec.EOF = False Then
'                nMaxCCodigo = IIf(IsNull(rec!Maximo), 1, rec!Maximo + 1)
'            End If
'            rec.Close
'            sql = "INSERT INTO CMEDICA"
'            sql = sql & " (CME_NUMERO,CME_FECHA, CLI_CODIGO,VEN_CODIGO,"
'            sql = sql & " MED_CODIGO,CME_INDICA)"
'            sql = sql & " VALUES ("
'            sql = sql & nMaxCCodigo & ","
'            sql = sql & XDQ(DTMedFec.Value) & ","
'            sql = sql & txtID & ","
'            sql = sql & cboMedDoc.ItemData(cboMedDoc.ListIndex) & ","
'            sql = sql & cboMedica.ItemData(cboMedica.ListIndex) & ","
'            sql = sql & XS(txtMedIndica.Text) & ")"
'
'
'        Else
'        ' Modifico Curso Clinico
'            If MsgBox("¿Confirma la Modificación de la Medicación?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'            ' aca hago el update
'            sql = "UPDATE CMEDICA SET "
'            sql = sql & " CME_FECHA = " & XDQ(DTMedFec.Value)
'            sql = sql & " ,CLI_CODIGO =" & XN(txtID)
'            sql = sql & " ,VEN_CODIGO = " & cboMedDoc.ItemData(cboMedDoc.ListIndex)
'            sql = sql & " ,MED_CODIGO = " & cboMedica.ItemData(cboMedica.ListIndex)
'            sql = sql & " ,CME_INDICA =" & XS(txtMedIndica.Text)
'            sql = sql & " WHERE CME_NUMERO = " & XN(txtMedCodigo.Text)
'
'
'        End If
'        DBConn.Execute sql
'        DBConn.CommitTrans
'        CargarCMedica txtID.Text
'        LimpiarCMedica
'
'    Exit Sub
'
'HayErrorCMedica:
'    Screen.MousePointer = vbNormal
'    If rec.State = 1 Then rec.Close
'    If Rec1.State = 1 Then Rec1.Close
'    DBConn.RollbackTrans
'    MsgBox Err.Description, vbCritical, TIT_MSGBOX
'
'End Sub
'
'Private Sub cmdMedNuevo_Click()
'    LimpiarCMedica
'End Sub
'
'Private Sub cmdMedQuitar_Click()
'    If txtMedCodigo.Text <> "" Then
'        If MsgBox("¿Elimina la Mediación?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'
'        On Error GoTo HayErrorCMedica
'        DBConn.BeginTrans
'
'        sql = "DELETE FROM CMEDICA WHERE CME_NUMERO =  " & XN(txtMedCodigo.Text)
'
'        DBConn.Execute sql
'        DBConn.CommitTrans
'        CargarCMedica txtID.Text
'        LimpiarCMedica
'    End If
'
'    Exit Sub
'
'HayErrorCMedica:
'    Screen.MousePointer = vbNormal
'    If rec.State = 1 Then rec.Close
'    If Rec1.State = 1 Then Rec1.Close
'    DBConn.RollbackTrans
'    MsgBox Err.Description, vbCritical, TIT_MSGBOX
'End Sub
'
'Private Sub CmdNuevo_Click()
'    LimpiarCClinico
'End Sub

'Private Sub cmdQuitar_Click()
'    If txtCCodigo.Text <> "" Then
'        If MsgBox("¿Elimina el Curso Clinico?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'
'        On Error GoTo HayErrorCClinico
'        DBConn.BeginTrans
'
'        sql = "DELETE FROM CCLINICO WHERE CCL_NUMERO =  " & XN(txtCCodigo.Text)
'
'        DBConn.Execute sql
'        DBConn.CommitTrans
'        CargarCClinico txtID.Text
'        LimpiarCClinico
'    End If
'
'    Exit Sub
'
'HayErrorCClinico:
'    Screen.MousePointer = vbNormal
'    If rec.State = 1 Then rec.Close
'    If Rec1.State = 1 Then Rec1.Close
'    DBConn.RollbackTrans
'    MsgBox Err.Description, vbCritical, TIT_MSGBOX
'
'End Sub
'
'Private Sub Command1_Click()
'    Dim X As Integer
'    X = 2
'    sql = "SELECT * FROM XX"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        Do While rec.EOF = False
'            sql = "INSERT INTO CLIENTE (CLI_CODIGO,CLI_RAZSOC,"
'            sql = sql & " CLI_DOMICI,CLI_TELEFONO,CLI_FAX,CLI_MAIL,CLI_CUMPLE,"
'            sql = sql & " IVA_CODIGO,PAI_CODIGO,PRO_CODIGO,LOC_CODIGO,CLI_NRODOC) VALUES ("
'            sql = sql & X & ","
'            sql = sql & "'" & Trim(rec!apellido) & " " & Trim(rec!Nombre) & "',"
'            sql = sql & XS(rec!DIRECCION) & ","
'            sql = sql & XS(rec!te) & ","
'            sql = sql & XS(rec!cel) & ","
'            sql = sql & XS(rec!mail) & ","
'            sql = sql & XDQ(ChkNull(rec!nacimiento)) & ",2,1,1,"
'            sql = sql & buscaloc(Trim(rec!CIUDAD)) & ","
'            sql = sql & XN(rec!DNI) & ")"
'            DBConn.Execute sql
'            X = X + 1
'            rec.MoveNext
'        Loop
'    End If
'End Sub

Private Function buscaloc(mlocdescri As String) As Integer
    Select Case mlocdescri
        Case "PILAR"
            buscaloc = 1
        Case "RIO SEGUNDO", "RIO II", "RIO 2", "RIO II  CBA"
            buscaloc = 2
        Case "COSTA SACATE"
            buscaloc = 6
        Case "LAGUNA LARGA"
            buscaloc = 5
        Case "LAGUNILLA"
            buscaloc = 10
        Case "ONCATIVO"
            buscaloc = 20
        Case "VILLA DEL ROSARIO"
            buscaloc = 7
        Case "TOLEDO"
            buscaloc = 3
        Case "LOZADA", "LOSADA"
            buscaloc = 4
        Case "MATORRALES"
            buscaloc = 17
        Case "DESPEÑADEROS"
            buscaloc = 9
        Case "IMPIRA"
            buscaloc = 21
        Case "CAPILLA DE LOS REMEDIOS"
            buscaloc = 25
        Case "CARLOS PAZ"
            buscaloc = 11
        Case "MINA CLAVERO"
            buscaloc = 12
        Case "CORDOBA"
            buscaloc = 13
        Case "VILLA DEL TOTORAL"
            buscaloc = 14
        Case "COSME SUD"
            buscaloc = 15
        Case "JAMES CRAIK"
            buscaloc = 19
        Case "PIQUILLIN"
            buscaloc = 18
        Case "LAS JUNTURAS"
            buscaloc = 22
        Case "CALCHIN OESTE"
            buscaloc = 24
        Case "RINCON"
            buscaloc = 8
        Case Else
            buscaloc = 1
    End Select
End Function

Private Sub cmdQuUno_Click()
    cmdAceptar.Enabled = True
    'If KeyCode = vbKeyDelete Then
        If grillaComCli.Rows = 2 Then
            grillaComCli.Rows = 1
        Else
            grillaComCli.RemoveItem (grillaComCli.RowSel)
        End If
    'End If
End Sub

Private Sub FecCanc_Change()
    cmdAceptar.Enabled = True
End Sub

'Private Sub chkAneste_Click()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub chkcardia_Click()
'    cmdAceptar.Enabled = True
'    If chkcardia.Value = Checked Then
'        AcCtrl txtcualca
'    Else
'        txtcualca.Text = ""
'        DesacCtrl txtcualca
'    End If
'End Sub
'
'Private Sub chkDiabet_Click()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub chkEmbara_Click()
'    cmdAceptar.Enabled = True
'    If chkEmbara.Value = Checked Then
'        AcCtrl txtMeses
'    Else
'        txtMeses.Text = ""
'        DesacCtrl txtMeses
'    End If
'End Sub
'
'Private Sub chkEpilep_Click()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub chkhemofi_Click()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub chkLactan_Click()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub chkmarcapaso_Click()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub chkprealt_Click()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub chkprebaj_Click()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub chktarcic_Click()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub chkTomaMed_Click()
'    cmdAceptar.Enabled = True
'    If chkTomaMed.Value = Checked Then
'        AcCtrl txtCualMe
'    Else
'        DesacCtrl txtCualMe
'    End If
'End Sub
'
'Private Sub chktuhemo_Click()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub DTFecha_Change()
'    'cmdAceptar.Enabled = True
'End Sub
'
'Private Sub DTFechaNac_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub DTFechaNac_LostFocus()
'     If Not IsNull(DTFechaNac) Then
'        txtEdad.Text = Calculo_Edad(DTFechaNac)
'     End If
'End Sub
'
'Private Sub DTFechaPCons_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub DTUltVis_Change()
'    cmdAceptar.Enabled = True
'End Sub
'Private Sub Form_Activate()
'    'hizo click en una columna no correcta
'    If vMode = 2 And vFieldID = "0" Then
'        Unload Me
'    End If
'
'
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub



Private Sub Form_Load()

    Dim cSQL As String
    Dim hSQL As String
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    
    'Me.Top = vFormLlama.Top + 1500
    'Me.Left = vFormLlama.Left + 1000
    Call Centrar_pantalla(Me)
    'CARGO COMBO CONDICIN IVA
    Call CargoComboBox(cboIva, "CONDICION_IVA", "IVA_CODIGO", "IVA_DESCRI")
    If cboIva.ListCount > 0 Then
        cboIva.ListIndex = 0
    End If
    
    'cargo el combo de PAIS
    cboPais.Clear
    cSQL = "SELECT * FROM PAIS WHERE PAI_CODIGO=1 ORDER BY PAI_DESCRI"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboPais.AddItem Trim(rec!PAI_DESCRI)
          cboPais.ItemData(cboPais.NewIndex) = rec!PAI_CODIGO
          rec.MoveNext
       Loop
       cboPais.ListIndex = cboPais.ListIndex + 1
    End If
    rec.Close
    cboPais_LostFocus
    
    configurogrilla
    
    CargarCboTipoCom
    CargarGrillaComida 0
    
    'LlenarComboDoctor
    'LlenarComboTratamiento
    DTFechaNac.Value = Null
    DTFechaPCons.Value = Null
    
    'DTFecha.Value = Date
    'DTMedFec.Value = Date
    'DTFecPC.Value = Null
    'LlenarComboMedic
    'anamnesis
    'DesacCtrl txtCualMe
    'DesacCtrl txtMeses
    'DesacCtrl txtcualca
    'DTUltVis.Value = Null
    
    Pais = ""
    Provincia = ""
    TabClientes.Tab = 0
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            If gPaciente <> 0 Then
                vMode = 2
                DesacCtrl txtID
                TabClientes.Tab = 2
                
                'Call BuscaCodigoProxItemData(frmTurnos.cboDoctor.ItemData(frmTurnos.cboDoctor.ListIndex), cboDoctor)
                'txtIdTra.Text = 1
                
                cSQL = "SELECT * FROM " & cTabla & "  WHERE CLI_CODIGO = " & gPaciente
            Else
                cSQL = "SELECT * FROM " & cTabla & "  WHERE CLI_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            End If
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                txtID.Text = rec!CLI_CODIGO
                txtNombre.Text = rec!CLI_RAZSOC
                
                Call BuscaCodigoProxItemData(rec!IVA_CODIGO, cboIva)
                txtCuit.Text = ChkNull(rec!CLI_CUIT)
                txtIngresosBrutos.Text = ChkNull(rec!CLI_INGBRU)
                DTFechaNac.Value = ChkNull(rec!CLI_CUMPLE)
                
                Call BuscaCodigoProxItemData(CInt(rec!PAI_CODIGO), cboPais)
                cboPais_LostFocus
                Pais = cboPais.Text
                
                Call BuscaCodigoProxItemData(CInt(rec!PRO_CODIGO), cboProvincia)
                cboProvincia_LostFocus
                Provincia = cboProvincia.Text
                
                txtNroDoc.Text = ChkNull(rec!CLI_NRODOC)
                Call BuscaCodigoProxItemData(CInt(rec!LOC_CODIGO), cboLocalidad)
                txtDNI.Text = ChkNull(rec!CLI_DNI)
                txtDomicilio.Text = ChkNull(rec!CLI_DOMICI)
                txtTelefono.Text = ChkNull(rec!CLI_TELEFONO)
                txtFax.Text = ChkNull(rec!CLI_FAX)
                txtCodPostal.Text = ChkNull(rec!CLI_CODPOS)
                txtMail.Text = ChkNull(rec!CLI_MAIL)
                txtObserva.Text = Trim(ChkNull(rec!CLI_OBSERVA))
                
                txtEdad.Text = ChkNull(rec!CLI_EDAD)
                txtOcupacion.Text = ChkNull(rec!CLI_OCUPACION)
                DTFechaPCons.Value = ChkNull(rec!CLI_FECPC)
                txtBuscaOS.Text = ChkNull(rec!OS_NUMERO)
                txtDiagnostico.Text = ChkNull(rec!CLI_DIAGNO)
                'fechaAlta.Value = Date
                fechaevento.Value = Date
                Select Case ChkNull(rec!CLI_FACTURA)
                    Case 1
                        optDiaria.Value = True
                    Case 2
                        optSemanal.Value = True
                    Case 3
                        optQuincenal.Value = True
                    Case 4
                        optMensual.Value = True
                End Select
                
                CargarProgramas
                CargarCancelados
                CargarComidas
                CargarViandas
                
                
                
                'txtBuscaOS_LostFocus
                'txtNAfiliado.Text = ChkNull(rec!CLI_NROAFIL)
                
'                txtMC.Text = ChkNull(rec!CLI_MC)
'                txtRelac.Text = ChkNull(rec!CLI_RELAC)
'                txtAFA.Text = ChkNull(rec!CLI_AFA)
'                txtAPP.Text = ChkNull(rec!CLI_APP)
'                txtEFisico.Text = ChkNull(rec!CLI_EFISICO)
'                txtDiag.Text = ChkNull(rec!CLI_DIAG)
'                txtEstCom.Text = ChkNull(rec!CLI_ESTCOM)
'                txtPTest.Text = ChkNull(rec!CLI_PTEST)
'                txtHC.Text = ChkNull(rec!CLI_HC)
'                txtMedica.Text = ChkNull(rec!CLI_MEDICA)
'
'                txtimagen.Text = ChkNull(rec!CLI_FOTO)
'                If txtimagen.Text <> "" Then
'                    Image1.Picture = LoadPicture(txtimagen.Text)
'                End If
'                txtAspCli.Text = ChkNull(rec!CLI_ASPCLI)
'                cargarAnamnesis
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    
    'establesco funcionalidad del form de datos
    SetMode vMode
    
    
    'CARGO EL CURSO CLINICO
'    If txtID.Text <> "" Then
        'CargarCClinico txtID.Text
        'CargarCMedica txtID.Text
'    End If
End Sub
Private Function CargarComidas()
    If gPaciente <> 0 Then
        cSQL = "SELECT C.COM_DESCRI,TC.TCOM_DESCRI,C.COM_CODIGO,C.TCOM_CODIGO "
        cSQL = cSQL & "FROM COMIDAS_CLIENTE CC,COMIDAS C, TIPO_COMIDA TC "
        cSQL = cSQL & " WHERE C.COM_CODIGO = CC.COM_CODIGO AND C.TCOM_CODIGO = TC.TCOM_CODIGO"
        cSQL = cSQL & " AND CLI_CODIGO = " & gPaciente
    Else
        cSQL = "SELECT C.COM_DESCRI,TC.TCOM_DESCRI,C.COM_CODIGO,C.TCOM_CODIGO "
        cSQL = cSQL & "FROM COMIDAS_CLIENTE CC,COMIDAS C, TIPO_COMIDA TC "
        cSQL = cSQL & " WHERE C.COM_CODIGO = CC.COM_CODIGO AND C.TCOM_CODIGO = TC.TCOM_CODIGO"
        cSQL = cSQL & " AND CLI_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
    End If
    
    Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (Rec1.BOF And Rec1.EOF) = 0 Then
        Do While Rec1.EOF = False
            grillaComCli.AddItem Rec1!COM_DESCRI & Chr(9) & Rec1!TCOM_DESCRI & Chr(9) & _
                                 Rec1!COM_CODIGO & Chr(9) & Rec1!TCOM_CODIGO
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Function

Private Function CargarCancelados()
    If gPaciente <> 0 Then
        cSQL = "SELECT * FROM CANCELA_CLIENTE WHERE CLI_CODIGO = " & gPaciente
    Else
        cSQL = "SELECT * FROM CANCELA_CLIENTE WHERE CLI_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
    End If
    
    Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (Rec1.BOF And Rec1.EOF) = 0 Then
        Do While Rec1.EOF = False
            grillaCan.AddItem Rec1!CAC_FECHA & Chr(9) & IIf(IsNull(Rec1!CAC_COMEN), "", Rec1!CAC_COMEN)
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Function
Private Function CargarViandas()
    If gPaciente <> 0 Then
        cSQL = "SELECT * FROM CLIENTE_VIANDAS WHERE CLI_CODIGO = " & gPaciente
    Else
        cSQL = "SELECT * FROM CLIENTE_VIANDAS WHERE CLI_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
    End If
    
    Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (Rec1.BOF And Rec1.EOF) = 0 Then
        Do While Rec1.EOF = False
            'grillaCan.AddItem Rec1!CAC_FECHA & Chr(9) & IIf(IsNull(Rec1!CAC_COMEN), "", Rec1!CAC_COMEN)
            chkVianda(Rec1!VIA_CODIGO - 1).Value = Checked
            'VER QUE HAGO CON EL PRECIO
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Function
Private Function CargarProgramas()
    If gPaciente <> 0 Then
        cSQL = "SELECT * FROM PROGRAMA_CLIENTE WHERE CLI_CODIGO = " & gPaciente
    Else
        cSQL = "SELECT * FROM PROGRAMA_CLIENTE WHERE CLI_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
    End If
    
    Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (Rec1.BOF And Rec1.EOF) = 0 Then
    
        Do While Rec1.EOF = False
            'ALMUERZO
            If Rec1!PRG_CODIGO < 9 Then
                If Rec1!PRG_CODIGO = 1 Then ' Domingo almuerzo
                    chkAlm(6).Value = Checked
                    txtAlm(6).Text = Rec1!PRC_CANT
                Else
                    If Rec1!PRG_CODIGO <> 0 Then
                        If Rec1!PRG_CODIGO = 8 Then ' feriado almuerzo
                            chkAlm(Rec1!PRG_CODIGO - 1).Value = Checked
                            txtAlm(Rec1!PRG_CODIGO - 1).Text = Rec1!PRC_CANT
                        Else
                            chkAlm(Rec1!PRG_CODIGO - 2).Value = Checked
                            txtAlm(Rec1!PRG_CODIGO - 2).Text = Rec1!PRC_CANT
                        End If
                    End If
                End If
            Else
                'CENA
                If Rec1!PRG_CODIGO = 9 Then ' Domingo almuerzo
                    chkCena(6).Value = Checked
                    txtCena(6).Text = Rec1!PRC_CANT
                Else
                    If Rec1!PRG_CODIGO = 16 Then ' feriado cena
                        chkCena(Rec1!PRG_CODIGO - 9).Value = Checked
                        txtCena(Rec1!PRG_CODIGO - 9).Text = Rec1!PRC_CANT
                    Else
                        chkCena(Rec1!PRG_CODIGO - 10).Value = Checked
                        txtCena(Rec1!PRG_CODIGO - 10).Text = Rec1!PRC_CANT
                    End If
                End If
            End If
        Rec1.MoveNext
        Loop
    
    End If
    Rec1.Close
End Function

'Private Function cargarAnamnesis()
'    If gPaciente <> 0 Then
'        cSQL = "SELECT * FROM CLIENTE_ANAM WHERE CLI_CODIGO = " & gPaciente
'    Else
'        cSQL = "SELECT * FROM CLIENTE_ANAM WHERE CLI_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
'    End If
'    Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
'    If (Rec1.BOF And Rec1.EOF) = 0 Then
'        'si encontró el registro muestro los datos
'        'txtID.Text = rec!CLI_CODIGO
'
'
'        chkTomaMed.Value = Chk0(Rec1!CLA_TOMMED)
'        If chkTomaMed.Value = Checked Then
'            AcCtrl txtCualMe
'        Else
'            DesacCtrl txtCualMe
'        End If
'        txtCualMe.Text = ChkNull(Rec1!CLA_CUALME)
'        txtAlergia.Text = ChkNull(Rec1!CLA_ALERGIA)
'        chkAneste.Value = Chk0(Rec1!CLA_ANESTE)
'        chktuhemo.Value = Chk0(Rec1!CLA_TUHEMO)
'        chktarcic.Value = Chk0(Rec1!CLA_TARCIC)
'        chkDiabet.Value = Chk0(Rec1!CLA_Diabet)
'        chkprealt.Value = Chk0(Rec1!CLA_prealt)
'        chkprebaj.Value = Chk0(Rec1!CLA_prebaj)
'        chkEpilep.Value = Chk0(Rec1!CLA_Epilep)
'        chkEmbara.Value = Chk0(Rec1!CLA_Embara)
'        txtMeses.Text = ChkNull(Rec1!CLA_MESES)
'        chkLactan.Value = Chk0(Rec1!CLA_Lactan)
'        chkhemofi.Value = Chk0(Rec1!CLA_hemofi)
'        chkcardia.Value = Chk0(Rec1!CLA_cardia)
'        txtcualca.Text = ChkNull(Rec1!CLA_CUALCA)
'        chkmarcapaso.Value = Chk0(Rec1!CLA_MARCAP)
''        DTUltVis.Value = Rec1!CLA_ULTVIS
'        If IsNull(Rec1!CLA_ULTVIS) Then
'            DTUltVis.Value = Null
'
'        Else
'            DTUltVis.Value = Rec1!CLA_ULTVIS
'        End If
'        If IsNull(Rec1!TR_CODIGO) Then
'            cboAnamTrat.ListIndex = 0
'        Else
'            Call BuscaCodigoProxItemData(Chk0(Rec1!TR_CODIGO), cboAnamTrat)
'        End If
'
'        txtcuadia.Text = ChkNull(Rec1!CLA_CUADIA)
'        txtAnamOtros.Text = ChkNull(Rec1!CLA_OTROS)
'    Else
'        'Beep
'        'MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
'    End If
'    Rec1.Close
'End Function
'Private Sub LlenarComboTratamiento()
'    sql = "SELECT * FROM TRATAMIENTO"
'    sql = sql & " ORDER BY TR_CODIGO"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        cboAnamTrat.AddItem " "
'        Do While rec.EOF = False
'            'cboTratamiento.AddItem rec!TR_DESCRI
'            'cboTratamiento.ItemData(cboTratamiento.NewIndex) = rec!TR_CODIGO
'            cboAnamTrat.AddItem rec!TR_DESCRI
'            cboAnamTrat.ItemData(cboAnamTrat.NewIndex) = rec!TR_CODIGO
'            rec.MoveNext
'        Loop
'        'cboTratamiento.ListIndex = -1
'        cboAnamTrat.ListIndex = 0
'    End If
'    rec.Close
'End Sub
'Private Sub LlenarComboDoctor()
'    sql = "SELECT * FROM VENDEDOR"
'    sql = sql & " WHERE PR_CODIGO =1"
'    sql = sql & " ORDER BY VEN_CODIGO"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        'cboFactura1.AddItem "(Todas)"
'        Do While rec.EOF = False
'            cboDoctor.AddItem rec!VEN_NOMBRE
'            cboDoctor.ItemData(cboDoctor.NewIndex) = rec!VEN_CODIGO
'            cboMedDoc.AddItem rec!VEN_NOMBRE
'            cboMedDoc.ItemData(cboMedDoc.NewIndex) = rec!VEN_CODIGO
'            rec.MoveNext
'        Loop
'        If User <> 99 Then
'            Call BuscaCodigoProxItemData(XN(User), cboDoctor)
'            Call BuscaCodigoProxItemData(XN(User), cboMedDoc)
'        Else
'            cboDoctor.ListIndex = -1
'            cboMedDoc.ListIndex = -1
'        End If
'    End If
'    rec.Close
'End Sub
'Private Sub LlenarComboMedic()
'    sql = "SELECT * FROM MEDICAMENTOS"
'    sql = sql & " ORDER BY MED_NOMBRE"
'    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'    If rec.EOF = False Then
'        'cboFactura1.AddItem "(Todas)"
'        Do While rec.EOF = False
'            cboMedica.AddItem rec!MED_NOMBRE
'            cboMedica.ItemData(cboMedica.NewIndex) = rec!MED_CODIGO
'            rec.MoveNext
'        Loop
'        cboMedica.ListIndex = -1
'    End If
'    rec.Close
'End Sub

Private Sub CargarCboTipoCom()
    cboTipoComida.Clear
    sql = "SELECT TCOM_CODIGO,TCOM_DESCRI"
    sql = sql & " FROM TIPO_COMIDA "
    sql = sql & " ORDER BY TCOM_DESCRI"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If (Rec1.BOF And Rec1.EOF) = 0 Then
          cboTipoComida.AddItem "<TODOS>"
       Do While Rec1.EOF = False
          cboTipoComida.AddItem Trim(Rec1!TCOM_DESCRI)
          cboTipoComida.ItemData(cboTipoComida.NewIndex) = Rec1!TCOM_CODIGO
          Rec1.MoveNext
       Loop
    End If
    cboTipoComida.ListIndex = 0
    Rec1.Close
End Sub
Private Sub CargarGrillaComida(Tipo As Integer)
    grillaComidas.Rows = 1
    cSQL = "SELECT C.COM_CODIGO,C.COM_DESCRI,TC.TCOM_DESCRI,TC.TCOM_CODIGO "
    cSQL = cSQL & "FROM COMIDAS C, TIPO_COMIDA TC "
    cSQL = cSQL & " WHERE C.TCOM_CODIGO = TC.TCOM_CODIGO"
    If Tipo <> 0 Then
        cSQL = cSQL & " AND C.TCOM_CODIGO = " & Tipo
    End If
    
    Rec1.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (Rec1.BOF And Rec1.EOF) = 0 Then
        Do While Rec1.EOF = False
            grillaComidas.AddItem Rec1!COM_DESCRI & Chr(9) & Rec1!TCOM_DESCRI & Chr(9) & _
                                 Rec1!COM_CODIGO & Chr(9) & Rec1!TCOM_CODIGO
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
    
End Sub

Private Sub configurogrilla()
    
    grillaComidas.FormatString = "Comida|Grupo|CodComida|TipoComida"
    grillaComidas.ColWidth(0) = 2000 'Comida
    grillaComidas.ColWidth(1) = 1600 'grupo
    grillaComidas.ColWidth(2) = 0 'CodComida
    grillaComidas.ColWidth(3) = 0 'TipoComida
    grillaComidas.Rows = 1
    grillaComidas.Cols = 4
    grillaComidas.BorderStyle = flexBorderNone
    grillaComidas.row = 0
    For i = 0 To grillaComidas.Cols - 1
        grillaComidas.Col = i
        grillaComidas.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grillaComidas.CellBackColor = &H808080    'GRIS OSCURO
        grillaComidas.CellFontBold = True
    Next
    
    grillaBaja.FormatString = "Fecha|Evento|Motivo"
    grillaBaja.ColWidth(0) = 1200 'Comida
    grillaBaja.ColWidth(1) = 1200 'grupo
    grillaBaja.ColWidth(2) = 1500 'CodComida
    
    grillaBaja.Rows = 1
    grillaBaja.Cols = 3
    grillaBaja.BorderStyle = flexBorderNone
    grillaBaja.row = 0
    For i = 0 To grillaBaja.Cols - 1
        grillaBaja.Col = i
        grillaBaja.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grillaBaja.CellBackColor = &H808080    'GRIS OSCURO
        grillaBaja.CellFontBold = True
    Next
    
    grillaComCli.FormatString = "Comida|Grupo|CodComida|TipoComida"
    grillaComCli.ColWidth(0) = 2000 'Comida
    grillaComCli.ColWidth(1) = 1600 'grupo
    grillaComCli.ColWidth(2) = 0 'CodComida
    grillaComCli.ColWidth(3) = 0 'TipoComida
    grillaComCli.Rows = 1
    grillaComCli.Cols = 4
    grillaComCli.BorderStyle = flexBorderNone
    grillaComCli.row = 0
    For i = 0 To grillaComCli.Cols - 1
        grillaComCli.Col = i
        grillaComCli.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grillaComCli.CellBackColor = &H808080    'GRIS OSCURO
        grillaComCli.CellFontBold = True
    Next
    
    
    'GRILLA CANCELADOS
    grillaCan.FormatString = "Fecha|Comentario"
    grillaCan.ColWidth(0) = 1200 'FECHA
    grillaCan.ColWidth(1) = 2000 'codigo para usar a futuro
    grillaCan.Rows = 1
    grillaCan.Cols = 2
    grillaCan.BorderStyle = flexBorderNone
    grillaCan.row = 0
    For i = 0 To grillaCan.Cols - 1
        grillaCan.Col = i
        grillaCan.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grillaCan.CellBackColor = &H808080    'GRIS OSCURO
        grillaCan.CellFontBold = True
    Next
    
    
    
    
    'medicamentos
'    GrdCMedica.FormatString = "^Fecha|<Medicamentos|<Indicaciones|Profesional Actuante|>Cod Trat|>cod Doc|nro hc"
'    GrdCMedica.ColWidth(0) = 1300 'FECHA
'    GrdCMedica.ColWidth(1) = 3000 'TRATAMIENTO
'    GrdCMedica.ColWidth(2) = 4000 'INDICACIONES
'    GrdCMedica.ColWidth(3) = 2500 'DOCTOR
'    GrdCMedica.ColWidth(4) = 0 'Codigo TRATAMIENTO
'    GrdCMedica.ColWidth(5) = 0 'CODIGO DOCTOR
'    GrdCMedica.ColWidth(6) = 0 'NRO HC
'    GrdCMedica.Rows = 1
'    GrdCMedica.Cols = 7
'    GrdCMedica.BorderStyle = flexBorderNone
'    GrdCMedica.row = 0
'    For i = 0 To GrdCMedica.Cols - 1
'        GrdCMedica.Col = i
'        GrdCMedica.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
'        GrdCMedica.CellBackColor = &H808080    'GRIS OSCURO
'        GrdCMedica.CellFontBold = True
'    Next

End Sub

Private Sub grillaBaja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If grillaBaja.Rows = 2 Then
            grillaBaja.Rows = 1
        Else
            grillaBaja.RemoveItem (grillaBaja.RowSel)
        End If
    End If
End Sub

Private Sub grillaCan_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdAceptar.Enabled = True
    If KeyCode = vbKeyDelete Then
        If grillaCan.Rows = 2 Then
            grillaCan.Rows = 1
        Else
            grillaCan.RemoveItem (grillaCan.RowSel)
        End If
    End If
End Sub

Private Sub grillaComCli_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdAceptar.Enabled = True
    If KeyCode = vbKeyDelete Then
        If grillaComCli.Rows = 2 Then
            grillaComCli.Rows = 1
        Else
            grillaComCli.RemoveItem (grillaComCli.RowSel)
        End If
    End If
End Sub

Private Sub txtAlm_Change(Index As Integer)
    cmdAceptar.Enabled = True
End Sub

Private Sub txtAlm_GotFocus(Index As Integer)
    SelecTexto txtAlm(Index)
End Sub

Private Sub txtAlm_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtAlm(Index), KeyAscii)
End Sub

Private Sub txtCena_Change(Index As Integer)
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCena_GotFocus(Index As Integer)
    SelecTexto txtCena(Index)
End Sub

Private Sub txtCena_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtCena(Index), KeyAscii)
End Sub

'Private Sub grdCClinico_Click()
'    If grdCClinico.Rows > 1 Then
'        DTFecha.Value = grdCClinico.TextMatrix(grdCClinico.RowSel, 0)
'        Call BuscaCodigoProxItemData(grdCClinico.TextMatrix(grdCClinico.RowSel, 5), cboDoctor)
'        'Call BuscaCodigoProxItemData(grdCClinico.TextMatrix(grdCClinico.RowSel, 4), cboTratamiento)
'        txtIdTra.Text = grdCClinico.TextMatrix(grdCClinico.RowSel, 4)
'        txtCodTra.Text = grdCClinico.TextMatrix(grdCClinico.RowSel, 8)
'        txtDescTra.Text = grdCClinico.TextMatrix(grdCClinico.RowSel, 1)
'        txtIndicaciones.Text = grdCClinico.TextMatrix(grdCClinico.RowSel, 2)
'        txtCCodigo.Text = grdCClinico.TextMatrix(grdCClinico.RowSel, 6)
'        nCCRowSel = grdCClinico.RowSel
'        DTFecPC.Value = grdCClinico.TextMatrix(grdCClinico.RowSel, 7)
'    End If
'End Sub
'
'Private Sub GrdCMedica_Click()
'    If GrdCMedica.Rows > 1 Then
'        DTMedFec.Value = GrdCMedica.TextMatrix(GrdCMedica.RowSel, 0)
'        Call BuscaCodigoProxItemData(GrdCMedica.TextMatrix(GrdCMedica.RowSel, 5), cboMedDoc)
'        Call BuscaCodigoProxItemData(GrdCMedica.TextMatrix(GrdCMedica.RowSel, 4), cboMedica)
'        txtMedIndica.Text = GrdCMedica.TextMatrix(GrdCMedica.RowSel, 2)
'        txtMedCodigo.Text = GrdCMedica.TextMatrix(GrdCMedica.RowSel, 6)
'        nCMRowSel = GrdCMedica.RowSel
'    End If
'End Sub



'Private Sub TabClientes_Click(PreviousTab As Integer)
'    If TabClientes.Tab <> 0 Then
'        Select Case TabClientes.Tab
'            Case 1
'                'chkTomaMed.SetFocus
'            Case 2
'            '    cboTratamiento.SetFocus
'            Case 3
'               ' cboMedica.SetFocus
'        End Select
'    Else
'        'txtDNI.SetFocus
'    End If
'End Sub

'Private Sub txtAFA_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtAlergia_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtAnamOtros_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtAPP_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtAspCli_Change()
'    'cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtBuscaOS_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtBuscaOS_GotFocus()
'    SelecTexto txtBuscaOS
'End Sub
'
'Private Sub txtBuscaOS_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF1 Then
'        BuscarOS "txtBuscaOS", "CODIGO"
'    End If
'End Sub

'Private Sub txtBuscaOS_KeyPress(KeyAscii As Integer)
'    KeyAscii = CarNumeroEntero(KeyAscii)
'End Sub
'
'Private Sub txtBuscaOS_LostFocus()
'    If txtBuscaOS.Text <> "" Then
'        cSQL = "SELECT OS_NUMERO, OS_NOMBRE FROM OBRA_SOCIAL WHERE OS_NUMERO = " & XN(txtBuscaOS.Text)
'        rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            txtBuscaOS.Text = ChkNull(rec!OS_NUMERO)
'            txtBuscarOSNombre.Text = ChkNull(rec!OS_NOMBRE)
'        Else
'            MsgBox "Obra Social inexistente", vbExclamation, TIT_MSGBOX
'            'txtBuscaOS.SetFocus
'        End If
'        rec.Close
'    End If
'End Sub

'Private Sub txtBuscarOSNombre_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtBuscarOSNombre_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF1 Then
'        BuscarOS "txtBuscaOS", "CODIGO"
'    End If
'End Sub
'
'Private Sub txtBuscarOSNombre_LostFocus()
'    If txtBuscaOS.Text = "" And txtBuscarOSNombre.Text <> "" Then
'        Set rec = New ADODB.Recordset
'        sql = "SELECT OS_NUMERO,OS_NOMBRE FROM OBRA_SOCIAL WHERE OS_NOMBRE LIKE '" & Trim(txtBuscarOSNombre.Text) & "%'"
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            If rec.RecordCount > 1 Then
'                BuscarOS "txtBuscaOS", "CADENA", Trim(txtBuscarOSNombre)
'                If rec.State = 1 Then rec.Close
'                txtBuscarOSNombre.SetFocus
'            Else
'                txtBuscaOS.Text = rec!OS_NUMERO
'                txtBuscarOSNombre.Text = ChkNull(rec!OS_NOMBRE)
'            End If
'
'        Else
'            If MsgBox("La Obra Social no existe!  ¿Desea agregarla al Sistema?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
'            'preguntar si quiere agregarlo y abrir abm de tratamientos
'            'MsgBox "Tratamiento inexistente", vbExclamation, TIT_MSGBOX
'                gObraS = 1
'                ABMObraSocial.txtDescri.Text = txtBuscarOSNombre.Text
'                ABMObraSocial.Show vbModal
'                txtBuscarOSNombre.SetFocus
'            Else
'                txtBuscaOS.SetFocus
'            End If
'            gObraS = 0
'        End If
'        If rec.State = 1 Then rec.Close
'    End If
'End Sub

Private Sub txtCodPostal_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCodPostal_GotFocus()
    SelecTexto txtCodPostal
End Sub

Private Sub txtCodPostal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

'Private Sub txtCodTra_Change()
'    If txtCodTra.Text = "" Then
'        txtDescTra.Text = ""
'    End If
'End Sub
'
'Private Sub txtCodTra_GotFocus()
'    SelecTexto txtCodTra
'End Sub
'
'Private Sub txtCodTra_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF1 Then
'        BuscarTratamientos "txtCodTra", "CODIGO"
'    End If
'End Sub
'
'Private Sub txtCodTra_KeyPress(KeyAscii As Integer)
'    KeyAscii = CarTexto(KeyAscii)
'End Sub
'
'Private Sub txtCodTra_LostFocus()
'    If txtCodTra.Text <> "" Then
'        sql = "SELECT TR_DESCRI,TR_CODIGO,TR_CODNUE FROM TRATAMIENTO WHERE TR_CODNUE LIKE '" & Trim(txtCodTra.Text) & "%'"
'        If rec.State = 1 Then rec.Close
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            If rec.RecordCount > 1 Then
'                BuscarTratamientos "txtCodTra", "CODIGO", Trim(txtCodTra)
'                If rec.State = 1 Then rec.Close
'                txtDescTra.SetFocus
'            Else
'                txtCodTra.Text = rec!TR_CODNUE
'                txtDescTra.Text = ChkNull(rec!TR_DESCRI)
'                txtIdTra.Text = rec!TR_CODIGO
'            End If
'        Else
'            MsgBox "Tratamiento inexistente", vbExclamation, TIT_MSGBOX
'            txtCodTra.SetFocus
'        End If
'        If rec.State = 1 Then rec.Close
'    End If
'End Sub
'
'Private Sub txtcuadia_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtcualca_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtCualMe_Change()
'    cmdAceptar.Enabled = True
'End Sub

Private Sub txtCuit_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCuit_GotFocus()
    SelecTexto txtCuit
End Sub

Private Sub txtCuit_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCuit_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(Trim(txtCuit.ClipText)) = 12 Then
      txtCuit.SelStart = 12
  End If
End Sub

Private Sub txtCuit_LostFocus()
    If txtCuit.Text <> "" Then
        'rutina de validación de CUIT
        If Not ValidoCuit(txtCuit) Then
            txtCuit.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtDiagnostico_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtDiagnostico_GotFocus()
    seltxt
End Sub

'Private Sub txtDescTra_Change()
'    If txtDescTra.Text = "" Then
'        txtCodTra.Text = ""
'    End If
'End Sub
'
'Private Sub txtDescTra_GotFocus()
'    SelecTexto txtDescTra
'End Sub
'
'Private Sub txtDescTra_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF1 Then
'        BuscarTratamientos "txtCodTra", "CODIGO"
'    End If
'End Sub
'
'Private Sub txtDescTra_LostFocus()
'    If txtCodTra.Text = "" And txtDescTra.Text <> "" Then
'        Set rec = New ADODB.Recordset
'        sql = "SELECT TR_CODNUE,TR_CODIGO,TR_DESCRI FROM TRATAMIENTO WHERE TR_DESCRI LIKE '" & Trim(txtDescTra.Text) & "%'"
'        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If rec.EOF = False Then
'            If rec.RecordCount > 1 Then
'                BuscarTratamientos "txtCodTra", "CADENA", Trim(txtDescTra)
'                If rec.State = 1 Then rec.Close
'                txtDescTra.SetFocus
'            Else
'                txtCodTra.Text = rec!TR_CODNUE
'                txtDescTra.Text = ChkNull(rec!TR_DESCRI)
'                txtIdTra.Text = rec!TR_CODIGO
'            End If
'
'        Else
'            If MsgBox("El Tratamiento no existe!  ¿Desea agregarlo al Sistema?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
'            'preguntar si quiere agregarlo y abrir abm de tratamientos
'            'MsgBox "Tratamiento inexistente", vbExclamation, TIT_MSGBOX
'                gTrata = 1
'                ABMTratamiento.txtDescri.Text = txtDescTra.Text
'                ABMTratamiento.Show vbModal
'                txtDescTra.SetFocus
'            Else
'                txtCodTra.SetFocus
'            End If
'            gTrata = 0
'        End If
'        If rec.State = 1 Then rec.Close
'    End If
'End Sub
'
'Private Sub txtDiag_Change()
'    cmdAceptar.Enabled = True
'End Sub

Private Sub txtDNI_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtDNI_GotFocus()
    SelecTexto txtDNI
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtDomCant_GotFocus()
    'SelecTexto txtDomCant
End Sub

Private Sub txtDomicilio_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtDomicilio_GotFocus()
    SelecTexto txtDomicilio
End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtEdad_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtEdad_GotFocus()
    SelecTexto txtEdad
End Sub

Private Sub txtEdad_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

'Private Sub txtEFisico_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtEstCom_Change()
'    cmdAceptar.Enabled = True
'End Sub

Private Sub txtFax_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtFax_GotFocus()
    SelecTexto txtFax
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtFerCant_GotFocus()
    'SelecTexto txtFerCant
End Sub

Private Sub txtImpVianda_LostFocus(Index As Integer)
    sumarVianda
End Sub

'Private Sub txtHC_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtIndicaciones_Change()
'    'cmdAceptar.Enabled = True
'End Sub

Private Sub txtIngresosBrutos_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtIngresosBrutos_GotFocus()
    SelecTexto txtIngresosBrutos
End Sub

Private Sub txtIngresosBrutos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtJueCant_GotFocus()
    'SelecTexto txtJueCant
End Sub

Private Sub txtLunCant_GotFocus()
    'SelecTexto txtLunCant
End Sub

Private Sub txtMail_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtMail_GotFocus()
    SelecTexto txtMail
End Sub

Private Sub txtMarCant_GotFocus()
    'SelecTexto txtMarCant
End Sub

Private Sub txtMieCant_GotFocus()
    'SelecTexto txtMieCant
End Sub

'Private Sub txtMC_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtMedica_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtMeses_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtNAfiliado_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtNAfiliado_KeyPress(KeyAscii As Integer)
'    KeyAscii = CarTexto(KeyAscii)
'End Sub

Private Sub txtNombre_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtNombre_GotFocus()
    seltxt
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtID_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtID_GotFocus()
    seltxt
End Sub

Private Sub txtID_LostFocus()

    Dim cSQL As String
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    
    If vMode = 1 Then ' si se esta usando en modo de nuevo registro
        If txtID.Text = "" Then
            If cSugerirID = True Then
                cSQL = "SELECT MAX(" & cCampoID & ") FROM " & cTabla
                'cSQL = cSQL & " WHERE PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
                rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
                If (rec.BOF And rec.EOF) = 0 Then
                    If rec.Fields(0) > 0 Then
                        txtID.Text = rec.Fields(0) + 1
                    Else
                        txtID.Text = 1
                    End If
                End If
            End If
        Else
            'verifico que no sea clave repetida
            cSQL = "SELECT COUNT(*) FROM " & cTabla & " WHERE " & cCampoID & " = " & XN(txtID.Text)
            'cSQL = cSQL & " AND PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                If rec.Fields(0) > 0 Then
                    Beep
                    MsgBox "Código de " & cDesRegistro & " repetido." & Chr(13) & _
                                     "El código ingresado Pertenece a otro registro de " & cDesRegistro & ".", vbCritical + vbOKOnly, App.Title
                    txtID.Text = ""
                    txtID.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub txtNroDoc_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtNroDoc_GotFocus()
    SelecTexto txtNroDoc
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtObserva_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtObserva_GotFocus()
    SelecTexto txtObserva
End Sub

Private Sub txtOcupacion_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtOcupacion_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtSabCant_GotFocus()
    'SelecTexto txtSabCant
End Sub

'Private Sub txtPTest_Change()
'    cmdAceptar.Enabled = True
'End Sub
'
'Private Sub txtRelac_Change()
'    cmdAceptar.Enabled = True
'End Sub

Private Sub txtTelefono_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtTelefono_GotFocus()
    SelecTexto txtTelefono
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

'Public Sub BuscarTratamientos(Txt As String, mQuien As String, Optional mCadena As String)
'    Dim cSQL As String
'    Dim hSQL As String
'    Dim B As CBusqueda
'    Dim i, posicion As Integer
'    Dim cadena As String
'
'    Set B = New CBusqueda
'    With B
'        cSQL = "SELECT TR_CODNUE,TR_DESCRI, TR_CODIGO"
'        cSQL = cSQL & " FROM TRATAMIENTO "
'        If mQuien = "CADENA" Then
'            cSQL = cSQL & " WHERE TR_DESCRI LIKE '" & Trim(mCadena) & "%'"
'        Else
'            If mCadena <> "" Then
'                cSQL = cSQL & " WHERE TR_CODNUE LIKE '" & Trim(mCadena) & "%'"
'            End If
'        End If
'
'        hSQL = "Codigo,Descripcion, Id"
'        .sql = cSQL
'        .Headers = hSQL
'        .Field = "TR_CODNUE"
'        campo1 = .Field
'        .Field = "TR_DESCRI"
'        campo2 = .Field
'        .Field = "TR_CODIGO"
'        campo3 = .Field
'
'        .OrderBy = "TR_CODNUE"
'        camponumerico = False
'        .Titulo = "Busqueda de Tratamientos :"
'        .MaxRecords = 1
'        .Show
'
'        ' utilizar la coleccion de datos devueltos
'        If .ResultFields.Count > 0 Then
'            If Txt = "txtCodTra" Then
'                txtCodTra.Text = .ResultFields(1)
'                txtCodTra_LostFocus
'            Else
'                'txtBuscaCliente.Text = .ResultFields(2)
'                'txtBuscaCliente_LostFocus
'            End If
'        End If
'    End With
'
'    Set B = Nothing
'
'End Sub
'Public Sub BuscarOS(Txt As String, mQuien As String, Optional mCadena As String)
'    Dim cSQL As String
'    Dim hSQL As String
'    Dim B As CBusqueda
'    Dim i, posicion As Integer
'    Dim cadena As String
'
'    Set B = New CBusqueda
'    With B
'        cSQL = "SELECT OS_NOMBRE, OS_NUMERO"
'        cSQL = cSQL & " FROM OBRA_SOCIAL "
'        If mQuien = "CADENA" Then
'            cSQL = cSQL & " WHERE OS_NOMBRE LIKE '" & Trim(mCadena) & "%'"
'        End If
'
'        hSQL = "Nombre, Código"
'        .sql = cSQL
'        .Headers = hSQL
'        .Field = "OS_NOMBRE"
'        campo1 = .Field
'        .Field = "OS_NUMERO"
'        campo2 = .Field
'
'        .OrderBy = "OS_NOMBRE"
'        camponumerico = False
'        .Titulo = "Busqueda de Obras Sociales :"
'        .MaxRecords = 1
'        .Show
'
'        ' utilizar la coleccion de datos devueltos
'        If .ResultFields.Count > 0 Then
'            If Txt = "txtBuscaOS" Then
'                txtBuscaOS.Text = .ResultFields(2)
'                txtBuscaOS_LostFocus
'            Else
'                'txtBuscaCliente.Text = .ResultFields(2)
'                'txtBuscaCliente_LostFocus
'            End If
'        End If
'    End With
'
'    Set B = Nothing
'
'End Sub

Private Sub txtVieCant_Change()

End Sub

Private Sub txtVieCant_GotFocus()
    'SelecTexto txtVieCant
End Sub


'End Function


