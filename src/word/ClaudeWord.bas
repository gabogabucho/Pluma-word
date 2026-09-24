Attribute VB_Name = "ClaudeWord"
' Claude para Word
' ----------------
' Envia el parrafo seleccionado a Claude Code y la respuesta aparece
' en stream debajo del parrafo. Todo ocurre dentro del documento:
'
'   - Una linea "@claude <instruccion>" dentro de la seleccion es la
'     instruccion (se borra del documento). Sin esa linea, Claude aplica
'     la tarea por defecto definida en el CLAUDE.md del workspace.
'   - La respuesta se escribe en un bloque protegido (control de
'     contenido bloqueado, en gris) hasta que termina.
'   - Ejecutar la macro con el cursor dentro de un bloque en curso
'     cancela la respuesta.
'
' La escritura en stream la hace ClaudeWord.exe (proceso oculto) via COM.
' Este archivo es ASCII a proposito: la importacion de VBA no soporta UTF-8,
' los acentos van con ChrW().

Option Explicit

Private Const TAG_PREFIX As String = "claude-"
Private Const INSTR_PREFIX As String = "@claude"
Private Const MAX_CHARS As Long = 60000
Private Const WD_COLOR_STATUS As Long = 15263976      ' #E8E8E8, casi blanco
Private Const WD_CC_HIDDEN As Long = 2                 ' wdContentControlHidden

Public Sub EnviarAClaude()
    On Error GoTo Fallo
    If Documents.Count = 0 Then Exit Sub

    Dim doc As Document
    Set doc = ActiveDocument

    If CancelarSiEstaEnCurso(doc) Then Exit Sub

    Dim instruccion As String, texto As String, linea As String
    Dim ancla As Range, p As Paragraph
    Dim aBorrar As New Collection

    For Each p In RangoAEnviar().Paragraphs
        linea = LimpiarLinea(p.Range.Text)
        If EsInstruccion(linea) Then
            instruccion = Trim$(Mid$(linea, Len(INSTR_PREFIX) + 1))
            If Left$(instruccion, 1) = ":" Then instruccion = Trim$(Mid$(instruccion, 2))
            aBorrar.Add p.Range
        ElseIf Len(linea) > 0 Then
            texto = texto & linea & vbLf
            Set ancla = p.Range
        End If
    Next p

    If ancla Is Nothing Then
        Application.StatusBar = "Claude: no hay texto para enviar."
        Exit Sub
    End If
    If Len(texto) > MAX_CHARS Then texto = Left$(texto, MAX_CHARS)

    Dim i As Long
    For i = aBorrar.Count To 1 Step -1
        aBorrar(i).Delete
    Next i

    ' Parrafo nuevo debajo del ancla, envuelto en un bloque protegido.
    ancla.InsertParagraphAfter
    Dim destino As Range
    Set destino = ancla.Paragraphs.Last.Range

    Dim id As String
    Randomize
    id = Format$(Now, "yyyymmddhhnnss") & CStr(Int(Rnd * 100000))

    Dim cc As ContentControl
    Set cc = doc.ContentControls.Add(wdContentControlRichText, destino)
    cc.Tag = TAG_PREFIX & id
    cc.Title = "Claude"
    cc.Range.Text = ChrW(10035) & " Claude est" & ChrW(225) & " pensando" & ChrW(8230)
    cc.Appearance = WD_CC_HIDDEN
    cc.Range.Font.Color = WD_COLOR_STATUS
    cc.LockContents = True

    doc.Variables.Add TAG_PREFIX & id & "-text", texto
    If Len(instruccion) > 0 Then doc.Variables.Add TAG_PREFIX & id & "-instr", instruccion

    Dim exe As String
    exe = Environ$("LOCALAPPDATA") & "\ClaudeWord\ClaudeWord.exe"
    If Len(Dir$(exe)) = 0 Then
        cc.LockContents = False
        cc.Range.Text = "Claude no est" & ChrW(225) & " instalado en esta PC."
        cc.Delete False
        Exit Sub
    End If

    Shell """" & exe & """ " & id, vbHide
    Application.StatusBar = "Claude est" & ChrW(225) & " escribiendo" & ChrW(8230)
    Exit Sub

Fallo:
    Application.StatusBar = "Claude: " & Err.Description
End Sub

' Si el cursor esta dentro de una respuesta en curso, la cancela.
' ClaudeWord.exe detecta que el bloque desaparecio y corta el proceso.
Private Function CancelarSiEstaEnCurso(doc As Document) As Boolean
    Dim cc As ContentControl
    For Each cc In doc.ContentControls
        If Left$(cc.Tag, Len(TAG_PREFIX)) = TAG_PREFIX Then
            If Selection.Range.InRange(cc.Range) Then
                cc.LockContents = False
                cc.Delete True
                Application.StatusBar = "Claude: respuesta cancelada."
                CancelarSiEstaEnCurso = True
                Exit Function
            End If
        End If
    Next cc
End Function

' Que se envia sin tener que seleccionar:
'   - Con seleccion: los parrafos seleccionados.
'   - Cursor en un parrafo normal: ese parrafo.
'   - Cursor en una linea "@claude ...": esa linea + el parrafo con texto de arriba.
Private Function RangoAEnviar() As Range
    If Selection.Type <> wdSelectionIP Then
        Set RangoAEnviar = Selection.Range
        Exit Function
    End If

    Dim r As Range
    Set r = Selection.Paragraphs(1).Range.Duplicate
    If EsInstruccion(LimpiarLinea(r.Text)) Then
        Dim anterior As Paragraph
        Set anterior = r.Paragraphs(1).Previous
        Do While Not anterior Is Nothing
            If Len(LimpiarLinea(anterior.Range.Text)) > 0 Then
                r.Start = anterior.Range.Start
                Exit Do
            End If
            Set anterior = anterior.Previous
        Loop
    End If
    Set RangoAEnviar = r
End Function

Private Function EsInstruccion(ByVal linea As String) As Boolean
    EsInstruccion = (LCase$(Left$(linea, Len(INSTR_PREFIX))) = INSTR_PREFIX)
End Function

Private Function LimpiarLinea(ByVal s As String) As String
    s = Replace(s, vbCr, "")
    s = Replace(s, Chr(7), "")
    s = Replace(s, Chr(11), " ")
    LimpiarLinea = Trim$(s)
End Function
