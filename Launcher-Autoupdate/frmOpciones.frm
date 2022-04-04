VERSION 5.00
Begin VB.Form frmOpciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Actualizador"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   4335
      Begin MDSLauncher.lvButtons_H lvButtons_H3 
         Height          =   495
         Left            =   3840
         TabIndex        =   10
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   873
         Caption         =   "?"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin MDSLauncher.lvButtons_H lvButtons_H2 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         Caption         =   "Resetear actualizaciones"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin MDSLauncher.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      Caption         =   "Aceptar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuración del juego"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CheckBox Check7 
         Caption         =   "Sincronización Vertical(Puede funcionar mal en algunas PC's)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   3975
      End
      Begin VB.CheckBox Check6 
         Caption         =   "AlphaBlending(Hechizos, meditaciones,etc)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   4095
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Efecto noche"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Limitar fps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Efectos de combate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Árboles con transparencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Caption         =   "No FullScreen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim settingFile As String
Private Sub Check1_Click()
    Call WriteVar(settingFile, "Init", "NoFullScreen", Check1.value)
End Sub

Private Sub Check2_Click()
    Call WriteVar(settingFile, "Init", "TreeTransparence", Check2.value)
    
End Sub

Private Sub Check3_Click()
    Call WriteVar(settingFile, "Init", "FightingEfects", Check3.value)
End Sub

Private Sub Check4_Click()
    Call WriteVar(settingFile, "Init", "FpsLimit", Check4.value)
End Sub

Private Sub Check5_Click()
    Call WriteVar(settingFile, "Init", "Night", Check5.value)
End Sub

Private Sub Check6_Click()
    Call WriteVar(settingFile, "Init", "AlphaBlending", Check6.value)
End Sub

Private Sub Check7_Click()
    Call WriteVar(settingFile, "Init", "VSync", Check7.value)
End Sub

Private Sub Form_Load()
    settingFile = App.Path & "/init/Settings.ini"
    Check1.value = Val(GetVar(settingFile, "Init", "NoFullScreen"))
    Check2.value = Val(GetVar(settingFile, "Init", "TreeTransparence"))
    Check3.value = Val(GetVar(settingFile, "Init", "FightingEfects"))
    Check4.value = Val(GetVar(settingFile, "Init", "FpsLimit"))
    Check5.value = Val(GetVar(settingFile, "Init", "Night"))
    Check6.value = Val(GetVar(settingFile, "Init", "AlphaBlending"))
    Check7.value = Val(GetVar(settingFile, "Init", "VSync"))
End Sub


 
Private Sub lvButtons_H1_Click()
    Unload Me
End Sub

Private Sub lvButtons_H2_Click()
    Call GuardarInt(App.Path & "\INIT\Update.ini", 0)
    Unload Me
    Call Main
End Sub

Private Sub lvButtons_H3_Click()
    MsgBox "En caso de algún error de actualizaciones(Versión obsoleta, errores en mapas, etc.), presionar el botón 'Resetear actualizaciones' puede solucionarlo"
End Sub
