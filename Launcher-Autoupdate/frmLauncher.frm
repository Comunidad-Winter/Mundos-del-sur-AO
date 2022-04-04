VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmLauncher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launcher MDSAO"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox recTxt 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6800
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmLauncher.frx":0000
   End
   Begin MDSLauncher.lvButtons_H cmd1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1085
      Caption         =   "Iniciar juego"
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
   Begin MDSLauncher.lvButtons_H cmd3 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1085
      Caption         =   "Salir"
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
   Begin MDSLauncher.lvButtons_H cmd2 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1085
      Caption         =   "Opciones"
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
   Begin MDSLauncher.lvButtons_H cmd4 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1085
      Caption         =   "Tengo errores al iniciar el juego"
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
   Begin VB.Label lblState 
      Alignment       =   2  'Center
      Caption         =   "¡¡3 actualizaciones encontradas!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   4575
   End
End
Attribute VB_Name = "frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&
Dim Directory As String, bDone As Boolean, dError As Boolean, f As Integer
Dim actualizando As Boolean
Dim numPatch As Byte


Private Sub Analizar()

    
    
    If HayUpdates Then
        actualizando = True
        cmd1.Enabled = False
        cmd2.Enabled = False
        cmd3.Enabled = False
        cmd4.Enabled = False
        For i = 1 To DifX
            Inet1.AccessType = icUseDefault
            dNum = i + tX
            lblState.ForeColor = RGB(100, 100, 100)
            numPatch = dNum
            lblState.Caption = "Descargando Parche" & numPatch & ".zip(0%)"
            
            Inet1.URL = "http://mundosdelsur.hol.es/AutoUpdate/Parche" & dNum & ".zip" 'Host
            
            
            Directory = App.Path & "\INIT\tmp.zip"
            bDone = False
            dError = False
                
            Inet1.Execute , "GET"
            
            Do While bDone = False
                DoEvents
            Loop
            
            If dError Then Exit Sub
            lblState.Caption = "Instalando Parche" & numPatch & ".zip"
            DoEvents
            UnZip Directory, App.Path & "\"
            DoEvents
            Kill Directory
            
        Next i
        Call GuardarInt(App.Path & "\INIT\Update.ini", iX)
        frmLauncher.lblState = "Cliente actualizado correctamente"
        frmLauncher.lblState.ForeColor = vbGreen
        DoEvents
        Sleep 3333
        cmd1.Enabled = True
        cmd2.Enabled = True
        cmd3.Enabled = True
        cmd4.Enabled = True
        actualizando = False
    End If
    Call GuardarInt(App.Path & "\INIT\Launcher.ini", 1)
    
    Call ShellExecute(Me.hwnd, "open", App.Path & "/MDSAO.exe", "", "", 1)
    End
End Sub

Private Sub cmd1_Click()
 Call Analizar
End Sub

Private Sub cmd2_Click()
    frmOpciones.Show vbModal, Me
End Sub

Private Sub cmd3_Click()
    Unload Me
End Sub

Sub AgregarLinea(ByRef Dest As String, ByVal Texto As String, ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, ByVal bold As Boolean, ByVal italic As Boolean)

End Sub

Private Sub cmd4_Click()
    Call ShellExecute(Me.hwnd, "open", App.Path & "/AdminLibSetup.exe", "", "", 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    Static Max As Long, value As Long
    Select Case State
        Case icError
            MsgBox "Hubo un error al actualizar el cliente."
           
            bDone = True
            dError = True
            End
        Case icResponseCompleted
            Dim vtData As Variant
            Dim tempArray() As Byte
            Dim FileSize As Long
            
            FileSize = Inet1.GetHeader("Content-length")
            Max = FileSize
            Open Directory For Binary Access Write As #1
                vtData = Inet1.GetChunk(1024, icByteArray)
                DoEvents
                
                
                Do While Not Len(vtData) = 0
                    tempArray = vtData
                    Put #1, , tempArray
                    
                    vtData = Inet1.GetChunk(1024, icByteArray)
                    
                    value = value + Len(vtData) * 2
                    ''LSize.Caption = (Value + Len(vtData) * 2) / 1000000 & "MBs de " & (Max / 1000000) & "MBs"
                    
                    ''ProgressBar1.Text = CLng((Value * 100) / Max) & "% Completado.]"
                    
                    lblState.Caption = "Descargando Parche" & numPatch & ".zip(" & Round((value * 100) / Max) & "%)"
                    DoEvents
                Loop
            Close #1
            ''LSize.Caption = FileSize & "bytes"
            value = 0
            
            bDone = True
    End Select
End Sub

