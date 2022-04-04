Attribute VB_Name = "modGral"
Option Explicit

Public HayUpdates As Boolean
Public iX As Integer, tX As Integer, DifX As Integer, dNum As String
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub CargarNoticias()
    
    Dim Texto As String, msg() As String, r As Byte, g As Byte, b As Byte
    ' ~34~177~76~1~0~1~1
    Texto = frmLauncher.Inet1.OpenURL("http://mundosdelsur.hol.es/AutoUpdate/Noticias.txt") ''"|Servidor Online~1| 02/05/2016 ~2 |En el dia de hoy se lanza la primera version de Mundos del Sur AO. Recordá que podes denunciar un bug o proponer una idea con /REPORTAR~0"
    msg = Split(Texto, "|")
    Dim X As Long
    frmLauncher.recTxt.TextRTF = ""
    For X = LBound(msg) To UBound(msg)
        AddToConsole frmLauncher.recTxt, msg(X)
    Next X
End Sub
Sub Main()
    If App.PrevInstance = True Then End
    frmLauncher.Show
    frmLauncher.cmd1.Enabled = False
    frmLauncher.cmd2.Enabled = False
    frmLauncher.cmd3.Enabled = False
    frmLauncher.cmd4.Enabled = False
    frmLauncher.lblState = "Buscando actualizaciones..."
    frmLauncher.lblState.ForeColor = RGB(100, 100, 100)
    
    
    iX = Val(frmLauncher.Inet1.OpenURL("http://mundosdelsur.hol.es/AutoUpdate/VEREXE.txt"))  'Host
    tX = LeerInt(App.Path & "\INIT\Update.ini")
    
    DifX = iX - tX
    
    CargarNoticias
    
    frmLauncher.cmd1.Enabled = True
    frmLauncher.cmd2.Enabled = True
    frmLauncher.cmd3.Enabled = True
    frmLauncher.cmd4.Enabled = True
    If DifX > 0 Then
        HayUpdates = True
        frmLauncher.lblState = "¡¡Se encontraron " & DifX & " nuevas actualizaciones!!"
        frmLauncher.lblState.ForeColor = vbRed
    Else
        HayUpdates = False
        frmLauncher.lblState = "El cliente está actualizado"
        frmLauncher.lblState.ForeColor = vbGreen
    End If
    
End Sub
Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Function LeerInt(ByVal Ruta As String) As Integer
    Dim f As Integer
    f = FreeFile
    If FileExist(Ruta, vbNormal) Then
        Open Ruta For Input As f
        LeerInt = Input$(LOF(f), #f)
        Close #f
    Else
        LeerInt = 0
    End If
End Function

Sub GuardarInt(ByVal Ruta As String, ByVal Data As Integer)
    Dim f As Integer
    f = FreeFile
    Open Ruta For Output As f
    Print #f, Data
    Close #f
End Sub
Sub AddToConsole(ByRef Rec As RichTextBox, ByVal Chat As String)
    Dim r As Byte, g As Byte, b As Byte, str As String, modo As Byte, bold As Boolean, italic As Boolean, underline As Boolean
    Dim newline As Boolean
    If InStr(1, Chat, "~") Then
        str = ReadField(2, Chat, Asc("~"))
        modo = Val(str)
        
        Select Case modo
            Case 0 'Cuerpo de la noticia
                r = 0
                g = 0
                b = 0
                bold = False
                italic = False
                underline = False
                newline = True
            Case 1 ''titulo
                ' ~34~177~76~1~0~1~1
                r = 34
                g = 177
                b = 76
                bold = True
                italic = False
                underline = True
                newline = True
            Case 2 ''fecha
                r = 100
                g = 100
                b = 100
                bold = False
                italic = False
                underline = False
                newline = False
        End Select
        
        Call AddtoRichTextBox(Rec, Left$(Chat, InStr(1, Chat, "~") - 1), r, g, b, bold, italic, newline, underline) '', Val(ReadField(6, Chat, Asc("~"))) <> 0, Val(ReadField(7, Chat, Asc("~"))) = 1, Val(ReadField(8, Chat, Asc("~"))) = 1)
        ''If modo = 0 Or modo = 2 Then Call AddtoRichTextBox(Rec, vbNullString, , , , , , True)
        If modo = 0 Then Call AddtoRichTextBox(Rec, vbNullString, , , , , , True)
    Else
        Call AddtoRichTextBox(Rec, Chat, 255, 255, 255, 1, 1)
    End If
End Sub


Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Chat As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True, Optional ByVal underline As Boolean)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************
    With RichTextBox
        
        
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        .SelUnderline = underline
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        If bCrLf And Len(.Text) > 0 Then Chat = vbCrLf & Chat
        .SelText = Chat
        
        RichTextBox.Refresh
    End With
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = Mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = Mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function
