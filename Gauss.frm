VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmGauss 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gauss Jordan - 2º Ano 4º Bimestre - Linguagem de Programação / Cálculo"
   ClientHeight    =   6870
   ClientLeft      =   1980
   ClientTop       =   975
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmGauss 
      Height          =   6855
      Left            =   0
      TabIndex        =   115
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton btnsobre 
         Caption         =   "Sobre"
         Height          =   255
         Left            =   7560
         TabIndex        =   138
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnlimpar 
         Caption         =   "Limpar"
         Height          =   495
         Left            =   7560
         TabIndex        =   136
         Top             =   6240
         Width           =   975
      End
      Begin VB.ListBox cmbExportar 
         Height          =   4155
         ItemData        =   "Gauss.frx":0000
         Left            =   7560
         List            =   "Gauss.frx":0002
         TabIndex        =   135
         Top             =   1920
         Width           =   975
      End
      Begin VB.HScrollBar HsbPasso 
         Height          =   255
         Left            =   120
         Max             =   1
         TabIndex        =   123
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton btnsair 
         Caption         =   "Sair"
         Height          =   255
         Left            =   7560
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton btnreset 
         Caption         =   "Reset"
         Height          =   255
         Left            =   7560
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton btnIniciar 
         Caption         =   "Iniciar"
         Height          =   255
         Left            =   7560
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton btnProcessar 
         Caption         =   "Processar "
         Height          =   495
         Left            =   120
         TabIndex        =   116
         Top             =   6240
         Width           =   7335
      End
      Begin VB.HScrollBar HsbMatriz 
         Height          =   255
         Left            =   6000
         Max             =   10
         Min             =   2
         TabIndex        =   1
         Top             =   480
         Value           =   2
         Width           =   1095
      End
      Begin TabDlg.SSTab TabGauss 
         Height          =   4815
         Left            =   120
         TabIndex        =   0
         Top             =   1320
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   8493
         _Version        =   393216
         Tabs            =   1
         TabHeight       =   520
         TabCaption(0)   =   "Matriz Principal"
         TabPicture(0)   =   "Gauss.frx":0004
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frmsis"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.Frame frmsis 
            Height          =   4215
            Left            =   120
            TabIndex        =   118
            Top             =   360
            Width           =   7095
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   110
               Left            =   6240
               TabIndex        =   114
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   109
               Left            =   5640
               TabIndex        =   113
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   108
               Left            =   5040
               TabIndex        =   112
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   107
               Left            =   4440
               TabIndex        =   111
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   106
               Left            =   3840
               TabIndex        =   110
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   105
               Left            =   3240
               TabIndex        =   109
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   104
               Left            =   2640
               TabIndex        =   108
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   103
               Left            =   2040
               TabIndex        =   107
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   102
               Left            =   1440
               TabIndex        =   106
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   101
               Left            =   840
               TabIndex        =   105
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   100
               Left            =   240
               TabIndex        =   104
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   99
               Left            =   6240
               TabIndex        =   103
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   98
               Left            =   5640
               TabIndex        =   102
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   97
               Left            =   5040
               TabIndex        =   101
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   96
               Left            =   4440
               TabIndex        =   100
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   95
               Left            =   3840
               TabIndex        =   99
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   94
               Left            =   3240
               TabIndex        =   98
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   93
               Left            =   2640
               TabIndex        =   97
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   92
               Left            =   2040
               TabIndex        =   96
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   91
               Left            =   1440
               TabIndex        =   95
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   90
               Left            =   840
               TabIndex        =   94
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   89
               Left            =   240
               TabIndex        =   93
               Top             =   3360
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   88
               Left            =   6240
               TabIndex        =   92
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   87
               Left            =   5640
               TabIndex        =   91
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   86
               Left            =   5040
               TabIndex        =   90
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   85
               Left            =   4440
               TabIndex        =   89
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   84
               Left            =   3840
               TabIndex        =   88
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   83
               Left            =   3240
               TabIndex        =   87
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   82
               Left            =   2640
               TabIndex        =   86
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   81
               Left            =   2040
               TabIndex        =   85
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   80
               Left            =   1440
               TabIndex        =   84
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   79
               Left            =   840
               TabIndex        =   83
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   78
               Left            =   240
               TabIndex        =   82
               Top             =   3000
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   77
               Left            =   6240
               TabIndex        =   81
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   76
               Left            =   5640
               TabIndex        =   80
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   75
               Left            =   5040
               TabIndex        =   79
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   74
               Left            =   4440
               TabIndex        =   78
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   73
               Left            =   3840
               TabIndex        =   77
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   72
               Left            =   3240
               TabIndex        =   76
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   71
               Left            =   2640
               TabIndex        =   75
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   70
               Left            =   2040
               TabIndex        =   74
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   69
               Left            =   1440
               TabIndex        =   73
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   68
               Left            =   840
               TabIndex        =   72
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   67
               Left            =   240
               TabIndex        =   71
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   66
               Left            =   6240
               TabIndex        =   70
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   65
               Left            =   5640
               TabIndex        =   69
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   64
               Left            =   5040
               TabIndex        =   68
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   63
               Left            =   4440
               TabIndex        =   67
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   62
               Left            =   3840
               TabIndex        =   66
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   61
               Left            =   3240
               TabIndex        =   65
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   60
               Left            =   2640
               TabIndex        =   64
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   59
               Left            =   2040
               TabIndex        =   63
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   58
               Left            =   1440
               TabIndex        =   62
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   57
               Left            =   840
               TabIndex        =   61
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   56
               Left            =   240
               TabIndex        =   60
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   55
               Left            =   6240
               TabIndex        =   59
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   54
               Left            =   5640
               TabIndex        =   58
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   53
               Left            =   5040
               TabIndex        =   57
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   52
               Left            =   4440
               TabIndex        =   56
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   51
               Left            =   3840
               TabIndex        =   55
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   50
               Left            =   3240
               TabIndex        =   54
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   49
               Left            =   2640
               TabIndex        =   53
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   48
               Left            =   2040
               TabIndex        =   52
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   47
               Left            =   1440
               TabIndex        =   51
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   46
               Left            =   840
               TabIndex        =   50
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   45
               Left            =   240
               TabIndex        =   49
               Top             =   1920
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   44
               Left            =   6240
               TabIndex        =   48
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   43
               Left            =   5640
               TabIndex        =   47
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   42
               Left            =   5040
               TabIndex        =   46
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   41
               Left            =   4440
               TabIndex        =   45
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   40
               Left            =   3840
               TabIndex        =   44
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   39
               Left            =   3240
               TabIndex        =   43
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   38
               Left            =   2640
               TabIndex        =   42
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   37
               Left            =   2040
               TabIndex        =   41
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   36
               Left            =   1440
               TabIndex        =   40
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   35
               Left            =   840
               TabIndex        =   39
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   34
               Left            =   240
               TabIndex        =   38
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   33
               Left            =   6240
               TabIndex        =   37
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   32
               Left            =   5640
               TabIndex        =   36
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   31
               Left            =   5040
               TabIndex        =   35
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   30
               Left            =   4440
               TabIndex        =   34
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   29
               Left            =   3840
               TabIndex        =   33
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   28
               Left            =   3240
               TabIndex        =   32
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   27
               Left            =   2640
               TabIndex        =   31
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   26
               Left            =   2040
               TabIndex        =   30
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   25
               Left            =   1440
               TabIndex        =   29
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   24
               Left            =   840
               TabIndex        =   28
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   23
               Left            =   240
               TabIndex        =   27
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   12
               Left            =   240
               TabIndex        =   16
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   13
               Left            =   840
               TabIndex        =   17
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   14
               Left            =   1440
               TabIndex        =   18
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   15
               Left            =   2040
               TabIndex        =   19
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   16
               Left            =   2640
               TabIndex        =   20
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   17
               Left            =   3240
               TabIndex        =   21
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   18
               Left            =   3840
               TabIndex        =   22
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   19
               Left            =   4440
               TabIndex        =   23
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   20
               Left            =   5040
               TabIndex        =   24
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   21
               Left            =   5640
               TabIndex        =   25
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   22
               Left            =   6240
               TabIndex        =   26
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   11
               Left            =   6240
               TabIndex        =   15
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   10
               Left            =   5640
               TabIndex        =   14
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   9
               Left            =   5040
               TabIndex        =   13
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   8
               Left            =   4440
               TabIndex        =   12
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   7
               Left            =   3840
               TabIndex        =   11
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   6
               Left            =   3240
               TabIndex        =   10
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   5
               Left            =   2640
               TabIndex        =   9
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   4
               Left            =   2040
               TabIndex        =   8
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   3
               Left            =   1440
               TabIndex        =   7
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   2
               Left            =   840
               TabIndex        =   6
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox valor 
               Height          =   285
               Index           =   1
               Left            =   240
               TabIndex        =   5
               Top             =   480
               Width           =   615
            End
            Begin VB.Label lblletra 
               Caption         =   "J"
               Height          =   255
               Index           =   9
               Left            =   5880
               TabIndex        =   134
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblletra 
               Caption         =   "I"
               Height          =   255
               Index           =   8
               Left            =   5280
               TabIndex        =   133
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblletra 
               Caption         =   "H"
               Height          =   255
               Index           =   7
               Left            =   4680
               TabIndex        =   132
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblletra 
               Caption         =   "G"
               Height          =   255
               Index           =   6
               Left            =   4080
               TabIndex        =   131
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblletra 
               Caption         =   "F"
               Height          =   255
               Index           =   5
               Left            =   3480
               TabIndex        =   130
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblletra 
               Caption         =   "E"
               Height          =   255
               Index           =   4
               Left            =   2880
               TabIndex        =   129
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblletra 
               Caption         =   "D"
               Height          =   255
               Index           =   3
               Left            =   2280
               TabIndex        =   128
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblletra 
               Caption         =   "C"
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   127
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblletra 
               Caption         =   "B"
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   126
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblletra 
               Caption         =   "A"
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   125
               Top             =   240
               Width           =   255
            End
         End
      End
      Begin VB.Label LblVer 
         Alignment       =   2  'Center
         Caption         =   "Matriz (X,Y)"
         Height          =   255
         Left            =   7560
         TabIndex        =   137
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label LblPas 
         Caption         =   "PASSO Nº"
         Height          =   255
         Left            =   120
         TabIndex        =   124
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblPasso 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   122
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblmatx 
         Caption         =   "x"
         Height          =   255
         Left            =   6840
         TabIndex        =   121
         Top             =   240
         Width           =   135
      End
      Begin VB.Label LblYval 
         Caption         =   "3"
         Height          =   255
         Left            =   6960
         TabIndex        =   120
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblXVal 
         Caption         =   "2"
         Height          =   255
         Left            =   6600
         TabIndex        =   119
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lbltamanho 
         Caption         =   "Matriz :"
         Height          =   255
         Left            =   6000
         TabIndex        =   117
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmGauss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mat(10, 11, 11) As Single
Dim X, Y, a, passo, pivo, i, J, k, piv(11) As Single

Private Sub btnlimpar_Click()
    cmbExportar.Clear
End Sub

Private Sub btnreset_Click()
    Form_Load
    HsbMatriz_Change
End Sub

Private Sub btnsair_Click()
    Unload Me
End Sub

Private Sub btnsobre_Click()
    MsgBox ("Enzo A. Marquiorato - Jamile G. de Oliveira - Wellington R. / FESP / 1999")
End Sub

Private Sub Form_Load()
    HsbMatriz.Enabled = True
    btnIniciar.Enabled = True
    btnreset.Enabled = False
    HsbPasso.Enabled = False
    btnProcessar.Enabled = False
    HsbPasso = 0
    HsbMatriz_Change
End Sub

Private Sub HsbMatriz_Change()
    X = HsbMatriz.Value
    Y = HsbMatriz.Value + 1
    lblXVal.Caption = X
    LblYval.Caption = HsbMatriz.Value
    HsbPasso.Max = X
    Montagem
End Sub
    
Private Sub Montagem()
    For J = 1 To 110
        valor(J).Visible = False
        valor(J).Enabled = False
        valor(J).Text = ""
    Next J
    For J = 1 To 9
        lblletra(J).Visible = False
    Next J
    a = 0
    For i = 1 To X
        For J = 1 To Y
            valor(J + a).Visible = True
            lblletra(i - 1).Visible = True
        Next J
        a = a + 11
    Next i
End Sub

Private Sub btnIniciar_Click()
    HsbMatriz.Enabled = False
    btnIniciar.Enabled = False
    btnreset.Enabled = True
    btnProcessar.Enabled = True
    a = 0
    For i = 1 To X
        For J = 1 To Y
            valor(J + a).Enabled = True
        Next J
        a = a + 11
    Next i
End Sub

Private Sub btnProcessar_Click()
    HsbPasso.Enabled = True
    HsbPasso.Value = 0
    For passo = 0 To X
        a = 0
        k = 0
        i = 1
        Do While i <= X
            J = 1
            Do While J <= Y
                If passo = 0 Then
                    mat(i, J, passo) = valor(J + a)
                End If
                If passo >= 1 Then
                    pivo = mat(passo, passo, (passo - 1))
                    f = 1
                    Do While pivo = 0
                        For e = 1 To Y
                            piv(e) = mat(passo, e, passo - 1)
                            mat(passo, e, passo - 1) = mat(passo + f, e, passo - 1)
                            mat(passo + f, e, passo - 1) = piv(e)
                        Next e
                        pivo = mat(passo, passo, passo - 1)
                        f = f + 1
                    Loop
                    If k = 0 Then
                       i = passo
                       mat(i, J, passo) = mat(i, J, (passo - 1)) / pivo
                       If J = Y Then
                            k = 1
                       End If
                    End If
                    If k > 1 Then
                        If i <> passo Then
                            mat(i, J, passo) = mat(i, J, passo - 1) - mat(i, passo, passo - 1) * mat(passo, J, passo)
                        End If
                    End If
               End If
               cmbExportar.AddItem mat(i, J, passo)
            J = J + 1
         Loop
        a = a + 11
        i = i + 1
        If k = 1 Then
            k = k + 1
            i = 1
        End If
        Loop
    Next passo
    HsbPasso_Change
End Sub

Private Sub HsbPasso_Change()
    For J = 1 To 110
        valor(J).Text = ""
    Next J
    lblPasso.Caption = HsbPasso.Value
    passo = HsbPasso.Value
    a = 0
    For i = 1 To X
        For J = 1 To Y
            valor(J + a).Text = mat(i, J, passo)
        Next J
        a = a + 11
    Next i
End Sub

