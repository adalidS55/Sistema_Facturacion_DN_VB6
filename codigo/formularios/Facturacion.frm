VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Facturacion 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura Contado"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   16710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardarPendiente 
      Caption         =   "Guardar pendiente / Nueva"
      Height          =   435
      Left            =   240
      TabIndex        =   54
      Top             =   6885
      Width           =   2895
   End
   Begin VB.ComboBox cboPendientes 
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   55
      Top             =   6945
      Width           =   10575
   End
   Begin VB.CommandButton cmdRecuperarPendiente 
      Caption         =   "Recuperar pendiente"
      Height          =   435
      Left            =   14160
      TabIndex        =   56
      Top             =   6885
      Width           =   2295
   End
   Begin VB.TextBox Text21 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """L""#,##0.0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   18442
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7920
      TabIndex        =   58
      Top             =   480
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Height          =   200
      Left            =   6120
      TabIndex        =   53
      Top             =   580
      Width           =   200
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   14520
      TabIndex        =   51
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   13680
      TabIndex        =   50
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   15120
      TabIndex        =   49
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   13920
      TabIndex        =   48
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "COBROS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   13680
      TabIndex        =   40
      Top             =   1080
      Width           =   2895
      Begin VB.TextBox Text17 
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   47
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   46
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text15 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   45
         Top             =   440
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "CAMBIO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1755
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "EFECTIVO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1155
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   515
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   15600
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AGREGAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16800
      TabIndex        =   35
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BORRAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16800
      TabIndex        =   34
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Total Ganacia"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   13455
      Begin VB.TextBox Text14 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   18442
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   14040
         TabIndex        =   38
         Top             =   4800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00FFC0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         TabIndex        =   36
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """L""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   18442
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9960
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   12360
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "Facturacion.frx":0000
         Left            =   2280
         List            =   "Facturacion.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox Text7 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """L""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   18442
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   14
         Top             =   4725
         Width           =   855
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H0080FF80&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """L""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   18442
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   13
         Top             =   4725
         Width           =   855
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H0080FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """L""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   18442
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5640
         TabIndex        =   12
         Top             =   4725
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   18442
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   12120
         TabIndex        =   11
         Top             =   4725
         Width           =   855
      End
      Begin VB.TextBox Text3 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """L""#,##0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   18442
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3615
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "T. Ganancia:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12720
         TabIndex        =   39
         Top             =   4800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Unit."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11280
         TabIndex        =   37
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cant."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   31
         Top             =   320
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   30
         Top             =   320
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   320
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Total"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "-- Desc."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   27
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "  =  Total  L."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   26
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. de Prod. Facturados"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9480
         TabIndex        =   25
         Top             =   4800
         Width           =   2655
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "P. Costo"
         Height          =   255
         Left            =   11400
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Desc."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   23
         Top             =   320
         Width           =   615
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Existencia"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   22
         Top             =   320
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8760
         TabIndex        =   21
         Top             =   320
         Width           =   855
      End
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   12000
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   120
      TabIndex        =   2
      Text            =   "CLIENTE CONTADO"
      Top             =   480
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   400
      Left            =   4320
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   155385858
      CurrentDate     =   43596
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   400
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   155385857
      CurrentDate     =   43596
   End
   Begin VB.Label lblPendientes 
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas pendientes: seleccione una y presione Recuperar."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   57
      Top             =   6540
      Width           =   16000
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "BORRAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15405
      TabIndex        =   59
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   705
      Left            =   15600
      Picture         =   "Facturacion.frx":0004
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   705
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "AGREGAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13920
      TabIndex        =   60
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   14115
      Picture         =   "Facturacion.frx":3499
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   705
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Descuento por cantidad:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6480
      TabIndex        =   52
      Top             =   373
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   7920
      TabIndex        =   41
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   705
      Left            =   14115
      Picture         =   "Facturacion.frx":5CA7
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   705
   End
   Begin VB.Image Image5 
      Height          =   705
      Left            =   15600
      Picture         =   "Facturacion.frx":93BE
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   705
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "LIMPIAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15465
      TabIndex        =   33
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "IMPRIMIR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13920
      TabIndex        =   32
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo Fact:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   10560
      TabIndex        =   6
      Top             =   500
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   0
      Picture         =   "Facturacion.frx":C8C5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "Facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private focoInicialAplicado As Boolean
Private cargandoPendiente As Boolean
Private rutasPendientes As New Collection
Private fila2 As Integer
Private inv As Double

Dim cn As New ADODB.Connection 'Creamos el objeto Connection.
Dim rs As New ADODB.Recordset  'Creamos el objeto Recordset.
Dim cn2 As New ADODB.Connection
Dim rs2 As New ADODB.Recordset
Dim cn3 As New ADODB.Connection
Dim rs3 As New ADODB.Recordset
Dim cn4 As New ADODB.Connection
Dim rs4 As New ADODB.Recordset

Private Sub Check1_Click()
If cargandoPendiente Then Exit Sub
If Check1.Value = Checked Then
Text21.Enabled = True
Text21.SetFocus
End If

If Check1.Value = Unchecked Then
Text21.Enabled = False
Text21.Text = ""
End If
End Sub


Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorProducto
If KeyCode = vbKeyReturn Then
    rs2.Requery
    If Len(Trim(Combo1.Text)) = 0 Then
        MsgBox "Debe Especificar el texto a buscar y digitar correctamente", vbExclamation, "Factura Contado"
        Combo1.SetFocus
        Exit Sub
    Else
        rs2.Find "nombreprod = '" & Replace$(Combo1.Text, "'", "''") & "'", , , 1
    End If
    If rs2.BOF = False And rs2.EOF = False Then
        Text1.Text = rs2.Fields("codproducto")
        Text4.Text = rs2.Fields("preciov")
        Text5.Text = CStr(ExistenciaVenta(cn, CStr(rs2!codproducto)))
        Text6.Text = rs2.Fields("precioc")

        Text2.SetFocus
    ElseIf rs2.BOF = True Or rs2.EOF = True Then
        MsgBox "El registro NO existe, digitar correctamente el codigo a buscar", vbExclamation, "Factura Contado"
        Combo1.Text = Empty
        Combo1.SetFocus
    End If
End If
   
    If KeyCode = vbKeyLeft Then
    Text1.SetFocus
    End If
    
    If KeyCode = vbKeyControl Then
    KeyCode = vbKeyF4
    End If
    
Exit Sub
ErrorProducto:
    MsgBox Err.Description, vbExclamation, "Facturacion"
End Sub

Private Sub Command1_Click()
Check1.Value = Unchecked
Dim i As Integer
If (Len(Trim$(Text1.Text)) = 0) Or (Combo1.Text = "") Or (Text2.Text = "" Or (Not (IsNumeric(Text2.Text)))) Or (Text3.Text = "" Or (Not (IsNumeric(Text3.Text)))) Or (Text4.Text = "" Or (Not (IsNumeric(Text4.Text)))) Or (Text5.Text = "" Or (Not (IsNumeric(Text5.Text)))) Or (Text6.Text = "" Or (Not (IsNumeric(Text6.Text)))) Then
    MsgBox "Datos en BLANCO, llenar las casillas", vbExclamation, "Factura Contado"
    Exit Sub
Else
    'For i = 0 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Col = 1
        MSFlexGrid1.Row = fila2
        MSFlexGrid1.Text = Text2.Text
        MSFlexGrid1.Col = 2
        MSFlexGrid1.Row = fila2
        MSFlexGrid1.Text = Text1.Text
        MSFlexGrid1.Col = 3
        MSFlexGrid1.Row = fila2
        MSFlexGrid1.Text = Combo1.Text
        MSFlexGrid1.Col = 4
        MSFlexGrid1.Row = fila2
        MSFlexGrid1.Text = Text4.Text
        MSFlexGrid1.Col = 5
        MSFlexGrid1.Row = fila2
        MSFlexGrid1.Text = Text3.Text
        MSFlexGrid1.Col = 6
        MSFlexGrid1.Row = fila2
        MSFlexGrid1.Text = (Val(Text2.Text) * Val(Text4.Text)) - (Val(Text3.Text) * Val(Text2.Text))
        MSFlexGrid1.Col = 7
        MSFlexGrid1.Row = fila2
        MSFlexGrid1.Text = Text6.Text
        MSFlexGrid1.Col = 8
        MSFlexGrid1.Row = fila2
        MSFlexGrid1.Text = ((Val(Text4.Text) - Val(Text6.Text)) - Val(Text3.Text)) * Val(Text2.Text)
        'descuento
        Text8.Text = Val(Text8) + (Val(MSFlexGrid1.TextMatrix(fila2, 5)) * Val(MSFlexGrid1.TextMatrix(fila2, 1)))
        'ganancia
        Text14.Text = Val(Text14) + (Val(MSFlexGrid1.TextMatrix(fila2, 8)))
        'total
        Text9.Text = Val(Text9.Text) + (Val(MSFlexGrid1.TextMatrix(fila2, 6)))
        'arregralsubtotal
        Text7.Text = Val(Text8) + Val(Text9)
        fila2 = fila2 + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    'Next
End If
borrartextofact
Text10.Text = MSFlexGrid1.Rows - 2

Text1.SetFocus
Text15.Text = Text9.Text
End Sub


Private Sub Command3_Click()
If MSFlexGrid1.Row <= 0 Then
    MsgBox "Debe Seleccionar una fila", vbInformation, "Factura Contado"
Else
    If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Empty Then
        MsgBox "No hay registro que eliminar", vbExclamation, "Factura Contado"
    ElseIf MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) <> Empty Then
        fila2 = fila2 - 1
        Text10.Text = Val(Text10.Text) - 1
        'descuento
        Text8.Text = Val(Text8.Text) - Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
        'ganancia
        Text14.Text = Val(Text14.Text) - Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8))
        'total
        Text9.Text = Val(Text9.Text) - Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
        Text7.Text = Val(Text9.Text) - Val(Text8.Text)
        Text15.Text = Val(Text9.Text)
        Text17.Text = Val(Text16.Text) - Val(Text15.Text)
        MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
    End If
End If
End Sub

Private Sub Command4_Click()
Text18.Text = (Val(Text19.Text) * Val(Text20.Text)) * 2
End Sub

''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  If TypeOf Me.ActiveControl Is MSFlexGrid Then FlexGridScroll Me.ActiveControl, MouseKeys, Rotation, Xpos, Ypos
End Sub


Private Sub Form_Activate()
    If focoInicialAplicado Then Exit Sub
    If Text1.Visible And Text1.Enabled Then
        Text1.SetFocus
        focoInicialAplicado = True
    End If
End Sub

Private Sub Form_Load()
    focoInicialAplicado = False

''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
  Dim N As Integer
  Dim i As Integer
  With MSFlexGrid1
    .Rows = 2
    .Cols = 9
    For N = .FixedRows To .Rows - 1
      .TextMatrix(N, 0) = "Row " & N
    Next N
  End With
  Call WheelHook(Me)
''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------


Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
    PrepararVinculaciones cn
rs.Source = "facturacion"
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Open "select * from facturacion", cn

Set rs2 = New ADODB.Recordset
cn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs2.Source = "productos"
rs2.CursorType = adOpenKeyset
rs2.LockType = adLockOptimistic
rs2.Open "select * from productos", cn

Set rs3 = New ADODB.Recordset
cn3.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs3.Source = "facturareporte"
rs3.CursorType = adOpenKeyset
rs3.LockType = adLockOptimistic
rs3.Open "select * from facturareporte", cn

Set rs4 = New ADODB.Recordset
cn4.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs4.Source = "codigosalternos"
rs4.CursorType = adOpenKeyset
rs4.LockType = adLockOptimistic
rs4.Open "select * from codigosalternos", cn4


While rs2.EOF = False
Combo1.AddItem rs2!nombreprod
rs2.MoveNext
Wend

fila2 = 1

cuadrofactura
ActualizarNumeroFactura
CargarPendientes

Image1.Left = 0
Image1.Top = 0
Image1.Height = Me.Height
Image1.Width = Me.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
cn2.Close
cn3.Close
cn4.Close

''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
Call WheelUnHook(Me)

End Sub

'Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Image2.Width > 705 Then
'Image2.Height = 705
'Image2.Width = 705
'Image2.Top = Image2.Top + 50
'Image2.Left = Image2.Left + 50
'End If
'End Sub

Private Sub Image2_Click()
Command1_Click
End Sub


'Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Image2.Width < 805 Then
'Image2.Height = 805
'Image2.Width = 805
'Image2.Top = Image2.Top - 50
'Image2.Left = Image2.Left - 50
'End If
'End Sub

Private Sub Image3_Click()
Command3_Click
End Sub


Private Sub Image4_Click()
Dim enTransaccion As Boolean
Dim mensaje As String
On Error GoTo ErrorGuardar
'Guardamos Factura
If (Text10.Text = "") Then
    MsgBox "No hay Datos en la FACTURA", vbInformation, "Factura contado"
    Exit Sub
Else
    
    If (Text10.Text <= 0) Or (Text10.Text = Empty) Or (Text9.Text <= 0) Or (Text9.Text = Empty) Or (MSFlexGrid1.TextMatrix(1, 2) = "") Then
        MsgBox "No hay Datos en la FACTURA", vbInformation, "Factura contado"
    Else
        Dim a As Integer
        Dim d As Integer
        Dim b As Integer
        Dim N As Integer
        Dim inv As Double
        ActualizarNumeroFactura
        cn.BeginTrans
        enTransaccion = True
        rs2.Requery
        N = MSFlexGrid1.Rows - 1
        For a = 1 To N - 1
            If (MSFlexGrid1.TextMatrix(a, 1) <> "") And (MSFlexGrid1.TextMatrix(a, 2) <> "") And (MSFlexGrid1.TextMatrix(a, 3) <> "") And (MSFlexGrid1.TextMatrix(a, 4) <> "") And (MSFlexGrid1.TextMatrix(a, 5) <> "") And (MSFlexGrid1.TextMatrix(a, 6) <> "") And (MSFlexGrid1.TextMatrix(a, 7) <> "") And (MSFlexGrid1.TextMatrix(a, 8) <> "") Then
                rs.AddNew
                rs("cantidad") = MSFlexGrid1.TextMatrix(a, 1)
                rs("codprodfact") = MSFlexGrid1.TextMatrix(a, 2)
                rs("nombreprodfact") = MSFlexGrid1.TextMatrix(a, 3)
                rs("precio") = MSFlexGrid1.TextMatrix(a, 4)
                rs("preciocosto") = MSFlexGrid1.TextMatrix(a, 7)
                rs("subtotal") = MSFlexGrid1.TextMatrix(a, 6)
                rs("gananciai") = MSFlexGrid1.TextMatrix(a, 8)
                rs("descuentoi") = MSFlexGrid1.TextMatrix(a, 5)
                
                rs("numfact") = Text12.Text
                rs("fechafact") = DTPicker1.Value
                rs("horafact") = DTPicker2.Value
                rs("subtotal") = Text7.Text
                rs("totaldescuento") = Text8.Text
                rs("totalneto") = Text9.Text
                rs("nprodfact") = Text10.Text
                rs("totalganacia") = Text14.Text
                rs.Update
    
            Else
                Err.Raise vbObjectError + 1000, , "Hay un producto incompleto en la factura."
            End If
        Next
        'Descontar del producto base segun la vinculacion y su factor fijo.
        For d = 1 To CLng(Text10.Text)
            ConsumirVenta cn, MSFlexGrid1.TextMatrix(d, 2), CCur(MSFlexGrid1.TextMatrix(d, 1))
        Next d
            'guardamos en reportes
            rs3.AddNew
                rs3("numfact") = Text12.Text
                rs3("fechafact") = DTPicker1.Value
                rs3("horafact") = DTPicker2.Value
                rs3("subtotal") = Text7.Text
                rs3("totaldescuento") = Text8.Text
                rs3("totalneto") = Text9.Text
                rs3("totalganacia") = Text14.Text
            rs3.Update
            cn.CommitTrans
            enTransaccion = False
            '********************
            MsgBox "Su Factura Ha Sido Registrada Exitosamente.", vbInformation, "Factura Contado"
            Image5_Click
            
            ActualizarNumeroFactura
    End If
End If
Exit Sub
ErrorGuardar:
    mensaje = Err.Description
    On Error Resume Next
    If rs.EditMode <> adEditNone Then rs.CancelUpdate
    If rs2.EditMode <> adEditNone Then rs2.CancelUpdate
    If rs3.EditMode <> adEditNone Then rs3.CancelUpdate
    If enTransaccion Then cn.RollbackTrans
    rs.Requery
    rs2.Requery
    rs3.Requery
    MsgBox "No se pudo guardar la factura: " & mensaje, vbExclamation, "Factura Contado"
End Sub

Private Sub Image5_Click()
NuevaFactura
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorProducto

Dim codigo_principal As String
If KeyCode = vbKeyReturn Then
    rs2.Requery
    rs4.Requery
End If
rs4.Find "Codalterno = '" & Replace$(Text1.Text, "'", "''") & "'", , , 1
If rs4.BOF = False And rs4.EOF = False Then
codigo_principal = rs4.Fields("Codprincipal")
Else
codigo_principal = Text1.Text
End If


If KeyCode = vbKeyReturn Then
    If Len(Trim(Text1.Text)) = 0 Then
        MsgBox "Debe Especificar el texto a buscar y digitar correctamente", vbExclamation, "Factura Contado"
        Text1.SetFocus
        Exit Sub
        
    Else
        rs2.Find "codproducto = '" & Replace$(codigo_principal, "'", "''") & "'", , , 1
    End If
    
    If rs2.BOF = False And rs2.EOF = False Then
        Text1.Text = codigo_principal
        Combo1.Text = rs2.Fields("nombreprod")
        Text4.Text = rs2.Fields("preciov")
        Text5.Text = CStr(ExistenciaVenta(cn, CStr(rs2!codproducto)))
        Text6.Text = rs2.Fields("precioc")
        Text2.SetFocus
    ElseIf rs2.BOF = True Or rs2.EOF = True Then
        MsgBox "El registro NO existe, digitar correctamente el codigo a buscar", vbExclamation, "Factura Contado"
        Text1.Text = Empty
        Text1.SetFocus
    End If

End If

    If KeyCode = vbKeyRight Then
    Combo1.SetFocus
    End If
    
Exit Sub
ErrorProducto:
    MsgBox Err.Description, vbExclamation, "Facturacion"
End Sub

Private Sub Text16_Change()
If cargandoPendiente Then Exit Sub
Text17.Text = Val(Text16.Text) - Val(Text15.Text)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorProducto
If KeyCode = vbKeyReturn Then
    If Text2.Text = "" Then
         MsgBox "Valores Vacios", vbCritical, "Factura Contado"
         Text2.SetFocus
         Exit Sub
    ElseIf (Not (IsNumeric(Text2.Text))) Then
        MsgBox "Solo Valores Numericos", vbCritical, "Factura Contado"
        Text2.SetFocus
        Exit Sub
    End If
' arreglaaaaarrrrrr
    If Text2.Text <= 0 Then
        MsgBox "Debe llenar el campo Corectamente", vbCritical, "Factura Contado"
        Text2.SetFocus
        Exit Sub
    Else
        Text13.Text = Val(Text2.Text) * Val(Text4.Text)
        inv = ExistenciaDisponibleFactura(Text1.Text) - CDbl(Text2.Text)
        If inv < 0 Then MsgBox "No existe suficiente inventario. Puede continuar con la venta; la existencia quedará en negativo.", vbExclamation, "Factura Contado"
        Text3.SetFocus
        Text3.Text = 0
    End If
End If

    If KeyCode = vbKeyRight Then
    Text3.SetFocus
    End If
    
    If KeyCode = vbKeyLeft Then
    Combo1.SetFocus
    End If

Exit Sub
ErrorProducto:
    MsgBox Err.Description, vbExclamation, "Facturacion"
End Sub

Private Sub Text21_Change()
If cargandoPendiente Then Exit Sub
If Len(Trim$(Text2.Text)) = 0 Then
MsgBox "Ingrese una cantidad primero", vbCritical, "Error Cantidad"
Text2.SetFocus
Else
If Text21.Text <> "" Then
Text3.Text = Text21.Text / Text2.Text
End If
End If

End Sub


Private Sub Text21_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Text3.SetFocus
Check1.Value = Unchecked
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Text3.Text = "" Then
         MsgBox "Valores NO validos. Si no se otorga decuento, favor colocar '0' (cero)", vbCritical, "Factura Contado"
         Text3.SetFocus
         Exit Sub
    ElseIf (Not (IsNumeric(Text3.Text))) Then
        MsgBox "Solo Valores Numericos", vbCritical, "Factura Contado"
        Text3.SetFocus
        Exit Sub
    End If
    If Text3.Text < 0 Then
        MsgBox "Debe llenar el campo Correctamente", vbCritical, "Factura Contado"
        Text3.SetFocus
        Exit Sub
    Else
        Text13.Text = (Val(Text2.Text) * Val(Text4.Text)) - Val(Text3.Text)
        If ((Val(Text4.Text) - Val(Text6.Text)) - Val(Text3.Text)) <= 0 Then
            MsgBox "El descuento otorgado es mayor que el margen de GANANCIA", vbCritical, "Factura Contado"
            Exit Sub
        Else
            Command1_Click
        End If
    End If
End If
    
    If KeyCode = vbKeyLeft Then
    Text2.SetFocus
    End If
    
End Sub

Private Sub Timer1_Timer()
DTPicker1.Value = Date
DTPicker2.Value = Time
End Sub

Private Sub borrartextofact()
Me.Text1 = Empty
Me.Text2 = Empty
Me.Combo1 = Empty
Me.Text3 = Empty
Me.Text4 = Empty
Me.Text5 = Empty
Me.Text6 = Empty
Me.Text13 = Empty
End Sub

Private Sub cuadrofactura()
Me.MSFlexGrid1.ColWidth(0) = 5
Me.MSFlexGrid1.ColWidth(1) = 800
Me.MSFlexGrid1.Col = 1
Me.MSFlexGrid1.Row = 0
Me.MSFlexGrid1.Text = "CANT."
Me.MSFlexGrid1.Col = 2
Me.MSFlexGrid1.Row = 0
Me.MSFlexGrid1.Text = "COD. PRODUCTO"
Me.MSFlexGrid1.ColWidth(2) = 2100
Me.MSFlexGrid1.ColAlignment(2) = 5
Me.MSFlexGrid1.Col = 3
Me.MSFlexGrid1.Row = 0
Me.MSFlexGrid1.Text = "D E T A L L E"
Me.MSFlexGrid1.ColWidth(3) = 6750
Me.MSFlexGrid1.ColAlignment(3) = 5
Me.MSFlexGrid1.Col = 4
Me.MSFlexGrid1.Row = 0
Me.MSFlexGrid1.Text = "PRECIO"
Me.MSFlexGrid1.ColWidth(4) = 1200
Me.MSFlexGrid1.Col = 5
Me.MSFlexGrid1.Row = 0
Me.MSFlexGrid1.Text = "DESC."
Me.MSFlexGrid1.ColWidth(5) = 1200
Me.MSFlexGrid1.Col = 6
Me.MSFlexGrid1.Row = 0
Me.MSFlexGrid1.Text = "TOTAL"
Me.MSFlexGrid1.ColWidth(6) = 700
Me.MSFlexGrid1.Col = 7
Me.MSFlexGrid1.Row = 0
'despues cambiar a 1
Me.MSFlexGrid1.ColWidth(7) = 1
Me.MSFlexGrid1.Text = "PRECIO COSTO"
'despues cambiar a 1
Me.MSFlexGrid1.ColWidth(8) = 1
Me.MSFlexGrid1.Col = 8
Me.MSFlexGrid1.Row = 0
Me.MSFlexGrid1.Text = "Ganancia individual"
End Sub

Private Sub limpiarcuadrofactura()
Dim a As Integer
Dim b As Integer
Me.MSFlexGrid1.Rows = 2
For a = 0 To Me.MSFlexGrid1.Rows - 1
  For b = 0 To Me.MSFlexGrid1.Cols - 1
  Me.MSFlexGrid1.TextMatrix(a, b) = ""
  Next
Next
End Sub


Private Sub ActualizarNumeroFactura()
    Dim numero As ADODB.Recordset
    Set numero = cn.Execute("SELECT Max(numfact) AS ultimo FROM facturacion")
    If IsNull(numero!ultimo) Then
        Text12.Text = "1"
    Else
        Text12.Text = CStr(CLng(numero!ultimo) + 1)
    End If
    numero.Close
    Set numero = Nothing
End Sub

Private Function CarpetaPendientes() As String
    CarpetaPendientes = rutaPendientes
    If Len(Dir$(CarpetaPendientes, vbDirectory)) = 0 Then MkDir CarpetaPendientes
End Function

Private Function TieneContenido() As Boolean
    Dim fila As Long, columna As Long
    'Ignorar encabezados, la columna auxiliar y las filas vacias del control.
    For fila = MSFlexGrid1.FixedRows To MSFlexGrid1.Rows - 1
        For columna = 1 To MSFlexGrid1.Cols - 1
            If Len(Trim$(MSFlexGrid1.TextMatrix(fila, columna))) > 0 Then
                TieneContenido = True
                Exit Function
            End If
        Next columna
    Next fila
End Function

Private Sub NuevaFactura()
    Dim i As Integer
    cargandoPendiente = True
    Check1.Value = vbUnchecked
    Text21.Enabled = False
    For i = 1 To 21
        If i <> 12 Then Controls("Text" & i).Text = ""
    Next i
    Text11.Text = "CLIENTE CONTADO"
    Combo1.Text = ""
    limpiarcuadrofactura
    cuadrofactura
    fila2 = 1
    cargandoPendiente = False
    ActualizarNumeroFactura
End Sub

Private Function EscribirPendiente() As String
    Dim pendiente As New FacturaPendiente
    Dim base As String, ruta As String, intento As Long
    If Not TieneContenido Then Exit Function
    base = CarpetaPendientes & "\" & Format$(Now, "dd-mm-yyyy hh-nn")
    Do
        intento = intento + 1
        ruta = base & "-" & Format$(intento, "000") & ".pend"
    Loop While Len(Dir$(ruta)) > 0 Or Len(Dir$(ruta & ".tmp")) > 0
    pendiente.Capturar Me
    pendiente.Guardar ruta
    EscribirPendiente = ruta
End Function

Private Sub CargarPendientes()
    Dim archivo As String, carpeta As String, ruta As Variant
    Dim archivos As New Collection, pendiente As FacturaPendiente
    carpeta = CarpetaPendientes
    archivo = Dir$(carpeta & "\*.pend")
    Do While Len(archivo) > 0
        archivos.Add carpeta & "\" & archivo
        archivo = Dir$()
    Loop
    cboPendientes.Clear
    Set rutasPendientes = New Collection
    For Each ruta In archivos
        Set pendiente = New FacturaPendiente
        On Error Resume Next
        pendiente.Leer CStr(ruta)
        If Err.Number = 0 Then
            cboPendientes.AddItem pendiente.Descripcion & " | " & NombreArchivoVisible(Mid$(CStr(ruta), Len(carpeta) + 2))
            rutasPendientes.Add CStr(ruta)
        Else
            cboPendientes.AddItem "No se puede leer: " & Mid$(CStr(ruta), Len(carpeta) + 2)
            rutasPendientes.Add CStr(ruta)
        End If
        Err.Clear
        On Error GoTo 0
    Next ruta
    If cboPendientes.ListCount > 0 Then cboPendientes.ListIndex = 0
    cmdRecuperarPendiente.Enabled = (cboPendientes.ListCount > 0)
    lblPendientes.Caption = "Pendientes: " & cboPendientes.ListCount & " - Al recuperar, la factura actual se guarda automaticamente si tiene productos en la tabla."
End Sub

Private Sub cmdGuardarPendiente_Click()
    On Error GoTo Fallo
    Dim ruta As String
    If Not TieneContenido Then
        MsgBox "Agregue al menos un producto a la tabla antes de guardar una pendiente.", vbInformation, "Facturas pendientes"
        Exit Sub
    End If
    ruta = EscribirPendiente
    NuevaFactura
    CargarPendientes
    Text1.SetFocus
    Exit Sub
Fallo:
    cargandoPendiente = False
    MsgBox "No se pudo guardar la pendiente: " & Err.Description, vbExclamation, "Facturas pendientes"
End Sub

Private Sub cmdRecuperarPendiente_Click()
    Dim pendiente As New FacturaPendiente, anterior As New FacturaPendiente
    Dim ruta As String, respaldo As String, mensaje As String
    Dim restaurando As Boolean
    On Error GoTo Fallo
    If cboPendientes.ListIndex < 0 Then Exit Sub
    ruta = rutasPendientes(cboPendientes.ListIndex + 1)
    pendiente.Leer ruta
    anterior.Capturar Me
    If TieneContenido Then respaldo = EscribirPendiente
    ActualizarNumeroFactura
    restaurando = True
    cargandoPendiente = True
    pendiente.Restaurar Me
    fila2 = MSFlexGrid1.Rows - 1
    cargandoPendiente = False
    'Solo retirar de la lista despues de restaurar todos los datos.
    Kill ruta
    restaurando = False
    CargarPendientes
    Text1.SetFocus
    Exit Sub
Fallo:
    mensaje = Err.Description
    On Error Resume Next
    If restaurando Then
        cargandoPendiente = True
        anterior.Restaurar Me
        fila2 = MSFlexGrid1.Rows - 1
    End If
    cargandoPendiente = False
    CargarPendientes
    MsgBox "No se pudo recuperar la pendiente: " & mensaje, vbExclamation, "Facturas pendientes"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ruta As String
    On Error GoTo Fallo
    If TieneContenido Then ruta = EscribirPendiente
    Exit Sub
Fallo:
    Cancel = 1
    MsgBox "No se pudo guardar la factura pendiente. La ventana seguira abierta: " & Err.Description, vbExclamation, "Facturas pendientes"
End Sub

Private Function ExistenciaDisponibleFactura(ByVal codigo As String) As Double
    Dim base As String, factor As Currency, otraBase As String, otroFactor As Currency
    Dim i As Long, disponible As Double
    ResolverInventario cn, codigo, base, factor
    disponible = ExistenciaVenta(cn, codigo)
    For i = 1 To MSFlexGrid1.Rows - 1
        If Len(Trim$(MSFlexGrid1.TextMatrix(i, 2))) > 0 Then
            ResolverInventario cn, MSFlexGrid1.TextMatrix(i, 2), otraBase, otroFactor
            If StrComp(base, otraBase, vbTextCompare) = 0 Then disponible = disponible - CDbl(MSFlexGrid1.TextMatrix(i, 1)) * CDbl(otroFactor) / CDbl(factor)
        End If
    Next i
    ExistenciaDisponibleFactura = disponible
End Function
