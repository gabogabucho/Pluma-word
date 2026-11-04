Attribute VB_Name = "ClaudeWord"
' Claude para Word
' ----------------
' Envia el parrafo seleccionado a Claude Code y la respuesta aparece
' en stream debajo del parrafo. Todo ocurre dentro del documento:
'
'   - Una linea "@claude <instruccion>" dentro de la seleccion es la
'     instruccion (se borra del documento). Sin esa linea, Claude aplica
'     la tarea por defecto definida en el CLAUDE.md del workspace.
'   - Las imagenes en linea del rango tambien se envian (el helper las extrae).
'   - La respuesta copia el formato del texto de ella (estilo, fuente, color).
'   - La respuesta se escribe en un bloque protegido (control de
'     contenido bloqueado, en gris) hasta que termina.
'   - Alt+Enter alterna: con una respuesta en curso la detiene (DetenerClaude)
'     y conserva lo ya escrito. Tambien desde el menu contextual.
'     (Alt+Shift+Enter no sirve: Windows usa Alt+Shift para cambiar el idioma.)
'
' La escritura en stream la hace ClaudeWord.exe (proceso oculto) via COM.
' Este archivo es ASCII a proposito: la importacion de VBA no soporta UTF-8,
' los acentos van con ChrW().

Option Explicit

Private Const TAG_PREFIX As String = "claude-"
Private Const INSTR_PREFIX As String = "@claude"
Private Const MAX_CHARS As Long = 60000
Private Const MAX_PARRAFOS As Long = 40               ' tope del bloque automatico
Private Const RESP_PREFIX As String = "_claude_r"     ' marcador de respuesta terminada
Private Const WD_COLOR_STATUS As Long = 15263976      ' #E8E8E8: se ve apenas, pero se ve
Private Const WD_CC_HIDDEN As Long = 2                 ' wdContentControlHidden

Private paso As String   ' ultimo paso de EnviarAClaude, para el log de errores

Public Sub EnviarAClaude()
    On Error GoTo Fallo
    If Documents.Count = 0 Then Exit Sub

    Dim doc As Document
    Set doc = ActiveDocument

    ' Alt+Enter alterna, como el boton de Claude: con una respuesta en curso, la detiene.
    paso = "detener"
    If HayEnCurso(doc) Then
        DetenerClaude
        Exit Sub
    End If

    Dim instruccion As String, texto As String, linea As String
    Dim ancla As Range, p As Paragraph, fuente As Range
    Dim aBorrar As New Collection
    Dim imagenes As Long

    paso = "leer seleccion"
    Set fuente = RangoAEnviar().Duplicate
    For Each p In fuente.Paragraphs
        linea = LimpiarLinea(p.Range.Text)
        If EsInstruccion(linea) Then
            instruccion = Trim$(Mid$(linea, Len(INSTR_PREFIX) + 1))
            If Left$(instruccion, 1) = ":" Then instruccion = Trim$(Mid$(instruccion, 2))
            aBorrar.Add p.Range
        Else
            ' Word representa cada imagen en linea con un caracter "/" en .Text.
            If p.Range.InlineShapes.Count > 0 Then
                imagenes = imagenes + p.Range.InlineShapes.Count
                If Len(Replace(Replace(linea, "/", ""), " ", "")) = 0 Then linea = ""
                Set ancla = p.Range
            End If
            If Len(linea) > 0 Then
                texto = texto & linea & vbLf
                Set ancla = p.Range
            End If
        End If
    Next p

    paso = "validar"
    If ancla Is Nothing Then
        Application.StatusBar = "Claude: no hay texto ni im" & ChrW(225) & "genes para enviar."
        LogMacro "nada para enviar (seleccion tipo " & Selection.Type & ")"
        Exit Sub
    End If
    If Len(texto) > MAX_CHARS Then texto = Left$(texto, MAX_CHARS)

    Dim i As Long
    paso = "borrar @claude"
    For i = aBorrar.Count To 1 Step -1
        aBorrar(i).Delete
    Next i

    Dim id As String
    Randomize
    id = Format$(Now, "yyyymmddhhnnss") & CStr(Int(Rnd * 100000))

    ' Con imagenes, el helper extrae el rango enviado via WordOpenXML.
    ' Marcador oculto (empieza con "_"): no aparece en el dialogo de marcadores.
    paso = "marcador"
    If imagenes > 0 Then doc.Bookmarks.Add "_claude_" & id, fuente

    ' Parrafo nuevo debajo del ancla, con el formato del texto de ella.
    paso = "parrafo nuevo"
    ancla.InsertParagraphAfter
    Dim destino As Range
    Set destino = ancla.Paragraphs.Last.Range
    paso = "copiar formato"
    CopiarFormato ancla, destino

    Dim cc As ContentControl
    paso = "bloque"
    Set cc = doc.ContentControls.Add(wdContentControlRichText, destino)
    cc.Tag = TAG_PREFIX & id
    cc.Title = "Claude"
    cc.Range.Text = ChrW(10035) & " Claude est" & ChrW(225) & " pensando" & ChrW(8230)
    cc.Appearance = WD_CC_HIDDEN
    cc.Range.Font.Color = WD_COLOR_STATUS
    cc.LockContents = True

    paso = "variables"
    If Len(texto) > 0 Then doc.Variables.Add TAG_PREFIX & id & "-text", texto
    If Len(instruccion) > 0 Then doc.Variables.Add TAG_PREFIX & id & "-instr", instruccion
    doc.Variables.Add TAG_PREFIX & id & "-color", CStr(ColorDelTexto(ancla))

    Dim exe As String
    exe = Environ$("LOCALAPPDATA") & "\ClaudeWord\ClaudeWord.exe"
    If Len(Dir$(exe)) = 0 Then
        cc.LockContents = False
        cc.Range.Text = "Claude no est" & ChrW(225) & " instalado en esta PC."
        cc.Delete False
        Exit Sub
    End If

    paso = "lanzar helper"
    Shell """" & exe & """ " & id, vbHide
    Application.StatusBar = "Claude est" & ChrW(225) & " escribiendo" & ChrW(8230)
    Exit Sub

Fallo:
    Application.StatusBar = "Claude: " & Err.Description
    LogMacro "error en '" & paso & "': " & Err.Number & " " & Err.Description
End Sub

' Errores de la macro en %LOCALAPPDATA%\ClaudeWord\logs\macro.log (la barra de
' estado se borra enseguida y es facil no verla).
Private Sub LogMacro(ByVal msg As String)
    On Error Resume Next
    Dim dir_ As String, f As Integer
    dir_ = Environ$("LOCALAPPDATA") & "\ClaudeWord\logs"
    If Len(Dir$(dir_, vbDirectory)) = 0 Then Exit Sub
    f = FreeFile
    Open dir_ & "\macro.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & msg
    Close #f
End Sub

' Detiene las respuestas en curso del documento activo y conserva lo ya escrito.
' Con el cursor dentro de un bloque, detiene solo ese; si no, todos.
Public Sub DetenerClaude()
    On Error GoTo Fallo
    If Documents.Count = 0 Then Exit Sub

    Dim doc As Document, cc As ContentControl, enCursor As String, n As Long
    Set doc = ActiveDocument
    For Each cc In doc.ContentControls
        If Left$(cc.Tag, Len(TAG_PREFIX)) = TAG_PREFIX Then
            n = n + 1
            If Selection.Range.InRange(cc.Range) Then enCursor = cc.Tag
        End If
    Next cc

    If n = 0 Then
        Application.StatusBar = "Claude: no hay ninguna respuesta en curso."
        Exit Sub
    End If
    For Each cc In doc.ContentControls
        If Left$(cc.Tag, Len(TAG_PREFIX)) = TAG_PREFIX Then
            If enCursor = "" Or cc.Tag = enCursor Then PedirParada doc, cc.Tag
        End If
    Next cc
    Application.StatusBar = "Claude: deteniendo" & ChrW(8230)
    Exit Sub

Fallo:
    Application.StatusBar = "Claude: " & Err.Description
    LogMacro "error al detener: " & Err.Number & " " & Err.Description
End Sub

Private Function HayEnCurso(doc As Document) As Boolean
    Dim cc As ContentControl
    For Each cc In doc.ContentControls
        If Left$(cc.Tag, Len(TAG_PREFIX)) = TAG_PREFIX Then
            HayEnCurso = True
            Exit Function
        End If
    Next cc
End Function

' Si el helper esta vivo, la macro no toca el bloque: le avisa con una variable
' del documento y el helper corta Claude y cierra el bloque ordenadamente.
' Si no hay helper (Word se cerro a mitad, o fallo), la macro cierra el bloque.
Private Sub PedirParada(doc As Document, ByVal tag As String)
    If HelperVivo(doc, tag) Then
        On Error Resume Next   ' si ya se pidio, la variable existe: no es un error
        doc.Variables.Add tag & "-stop", "1"
    Else
        CerrarHuerfano doc, tag
    End If
End Sub

' El helper actualiza "<tag>-hb" (fecha OLE) cada segundo mientras trabaja.
Private Function HelperVivo(doc As Document, ByVal tag As String) As Boolean
    On Error GoTo NoVivo
    Dim hb As Double
    hb = Val(doc.Variables(tag & "-hb").Value)   ' Val usa "." siempre
    HelperVivo = (CDbl(Now) - hb) * 86400# < 6
    Exit Function
NoVivo:
    HelperVivo = False
End Function

' Cierra un bloque sin helper: conserva lo escrito (marcado con " ...") o, si
' todavia decia "pensando", lo quita junto con su renglon vacio.
Private Sub CerrarHuerfano(doc As Document, ByVal tag As String)
    On Error Resume Next
    Dim ccs As ContentControls, cc As ContentControl, par As Range
    Set ccs = doc.SelectContentControlsByTag(tag)
    If ccs.Count > 0 Then
        Set cc = ccs(1)
        cc.LockContents = False
        If Left$(cc.Range.Text, 1) = ChrW(10035) Then
            Set par = cc.Range.Paragraphs(1).Range
            cc.Delete True
            If Len(LimpiarLinea(par.Text)) = 0 Then par.Delete
        Else
            cc.Range.InsertAfter " " & ChrW(8230)
            cc.Delete False
        End If
    End If
    Dim sufijo As Variant
    For Each sufijo In Array("-text", "-instr", "-color", "-stop", "-hb")
        doc.Variables(tag & sufijo).Delete
    Next sufijo
    doc.Bookmarks("_claude_" & Mid$(tag, Len(TAG_PREFIX) + 1)).Delete
End Sub

' La respuesta hereda estilo, fuente y parrafo del texto de ella: InsertParagraphAfter
' solo copia el formato de la marca de parrafo, que puede no coincidir con el de las
' letras. Si ella escribio un titulo o un item de lista, la respuesta va como Normal.
Private Sub CopiarFormato(ByVal origen As Range, ByVal destino As Range)
    Dim par As Paragraph
    Set par = origen.Paragraphs(1)
    If par.OutlineLevel <> wdOutlineLevelBodyText Then
        destino.Style = wdStyleNormal
        destino.ListFormat.RemoveNumbers
        Exit Sub
    End If
    destino.Style = par.Style
    destino.ParagraphFormat = par.Format
    destino.ListFormat.RemoveNumbers
    destino.Font = PrimerCaracterDeTexto(origen).Font.Duplicate
End Sub

' Primer caracter visible (no imagen ni espacio), para copiar su fuente.
Private Function PrimerCaracterDeTexto(ByVal r As Range) As Range
    Dim c As Range
    For Each c In r.Characters
        If c.InlineShapes.Count = 0 And Len(Trim$(LimpiarLinea(c.Text))) > 0 Then
            Set PrimerCaracterDeTexto = c
            Exit Function
        End If
    Next c
    Set PrimerCaracterDeTexto = r.Characters(1)
End Function

' Color del texto de ella (wdColorAutomatic si es mixto o no hay texto).
Private Function ColorDelTexto(ByVal r As Range) As Long
    Dim color As Long
    color = PrimerCaracterDeTexto(r).Font.Color
    If color = wdUndefined Then color = wdColorAutomatic
    ColorDelTexto = color
End Function

' Que se envia sin tener que seleccionar. Mensaje = BLOQUE: los renglones
' seguidos con contenido (texto o imagen), cortando en lineas vacias y en
' respuestas anteriores de Claude.
'   - Con seleccion: exactamente lo seleccionado.
'   - Cursor en un bloque: todo el bloque.
'   - Cursor en una linea vacia: el bloque de arriba.
'   - Bloque con solo lineas "@claude ...": se suma el bloque de arriba.
Private Function RangoAEnviar() As Range
    If Selection.Type <> wdSelectionIP Then
        Set RangoAEnviar = Selection.Range
        Exit Function
    End If

    ' Los marcadores de respuesta son ocultos: sin ShowHidden no se ven.
    Dim doc As Document, ocultos As Boolean
    Set doc = Selection.Document
    ocultos = doc.Bookmarks.ShowHidden
    doc.Bookmarks.ShowHidden = True
    On Error GoTo Restaurar

    Dim p As Paragraph
    Set p = Selection.Paragraphs(1)
    Do While EsVacio(p)
        If p.Previous Is Nothing Then Exit Do
        Set p = p.Previous
    Loop

    Dim r As Range
    Set r = BloqueDe(p)
    If SoloInstrucciones(r) Then
        Dim arriba As Paragraph
        Set arriba = r.Paragraphs(1).Previous
        Do Until arriba Is Nothing
            If Not EsVacio(arriba) Then Exit Do
            Set arriba = arriba.Previous
        Loop
        If Not arriba Is Nothing Then r.Start = BloqueDe(arriba).Start
    End If
    Set RangoAEnviar = r

Restaurar:
    doc.Bookmarks.ShowHidden = ocultos
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Function

' Renglones seguidos alrededor de p. Corta en lineas vacias, en el limite de
' parrafos y donde cambia "respuesta de Claude / texto de ella".
Private Function BloqueDe(ByVal p As Paragraph) As Range
    Dim r As Range, q As Paragraph, n As Long, resp As Boolean
    Set r = p.Range.Duplicate
    If EsVacio(p) Then
        Set BloqueDe = r
        Exit Function
    End If
    resp = EsRespuesta(p)

    Set q = p.Previous
    Do Until q Is Nothing
        If n >= MAX_PARRAFOS Then Exit Do
        If EsVacio(q) Then Exit Do
        If EsRespuesta(q) <> resp Then Exit Do
        r.Start = q.Range.Start
        n = n + 1
        Set q = q.Previous
    Loop

    Set q = p.Next
    Do Until q Is Nothing
        If n >= MAX_PARRAFOS Then Exit Do
        If EsVacio(q) Then Exit Do
        If EsRespuesta(q) <> resp Then Exit Do
        r.End = q.Range.End
        n = n + 1
        Set q = q.Next
    Loop
    Set BloqueDe = r
End Function

Private Function EsVacio(ByVal q As Paragraph) As Boolean
    EsVacio = (Len(LimpiarLinea(q.Range.Text)) = 0 And q.Range.InlineShapes.Count = 0)
End Function

' El helper marca cada respuesta terminada con un marcador oculto "_claude_r<id>".
Private Function EsRespuesta(ByVal q As Paragraph) As Boolean
    Dim b As Bookmark
    For Each b In q.Range.Bookmarks
        If Left$(b.Name, Len(RESP_PREFIX)) = RESP_PREFIX Then
            EsRespuesta = True
            Exit Function
        End If
    Next b
End Function

Private Function SoloInstrucciones(ByVal r As Range) As Boolean
    Dim q As Paragraph
    For Each q In r.Paragraphs
        If Not EsVacio(q) And Not EsInstruccion(LimpiarLinea(q.Range.Text)) Then Exit Function
    Next q
    SoloInstrucciones = True
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
