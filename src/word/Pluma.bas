Attribute VB_Name = "Pluma"
' Pluma para Word
' ---------------
' Envia el parrafo al agente elegido y la respuesta aparece en stream
' debajo. Todo ocurre dentro del documento:
'
'   - Una linea "@agente <instruccion>" elige el agente y, si hay texto
'     despues, esa es la instruccion (la linea se borra). Sin esa linea
'     se usa el agente de settings.json (Configuracion) y, si no hay,
'     el de agents.json.
'   - Las tareas del menu (corregir, formal, resumir, expandir, traducir)
'     mandan su id; el helper lee el texto en tasks.json.
'   - Las imagenes en linea del rango tambien se envian si el agente puede.
'   - La respuesta copia el formato del texto (estilo, fuente, color).
'   - La respuesta se escribe en un bloque protegido hasta que termina.
'   - El atajo (de fabrica Alt+Enter) alterna: con una respuesta en
'     curso la detiene. Sale de settings.json. Alt+Shift lo usa Windows
'     para cambiar el idioma, asi que no llega a Word.
'
' La escritura en stream la hace Pluma.exe (proceso oculto) via COM.
' Este archivo es ASCII a proposito: la importacion de VBA no soporta UTF-8.

Option Explicit

Private Const TAG_PREFIX As String = "pluma-"
Private Const MAX_CHARS As Long = 60000
Private Const MAX_PARRAFOS As Long = 40
Private Const RESP_PREFIX As String = "_pluma_r"
Private Const WD_COLOR_STATUS As Long = 15263976
Private Const WD_CC_HIDDEN As Long = 2

Private paso As String

Public Sub EnviarAPluma()
    Enviar "", True
End Sub

Public Sub TareaCorregir()
    Enviar "corregir", False
End Sub

Public Sub TareaFormal()
    Enviar "formal", False
End Sub

Public Sub TareaResumir()
    Enviar "resumir", False
End Sub

Public Sub TareaExpandir()
    Enviar "expandir", False
End Sub

Public Sub TareaTraducir()
    Enviar "traducir", False
End Sub

Public Sub ConfigurarPluma()
    On Error GoTo Fallo
    Dim exe As String
    exe = Environ$("LOCALAPPDATA") & "\Pluma\Pluma.exe"
    If Len(Dir$(exe)) = 0 Then
        Application.StatusBar = "Pluma no est" & ChrW(225) & " instalado en esta PC."
        Exit Sub
    End If
    Shell """" & exe & """ --config", vbNormalFocus
    Exit Sub
Fallo:
    Application.StatusBar = "Pluma: " & Err.Description
End Sub

' Al abrir Word, y cuando la ventana de configuracion lo pide.
Public Sub AutoExec()
    AplicarAjustesPluma
End Sub

Public Sub AplicarAjustesPluma()
    On Error GoTo Fin
    Dim carpeta As String, json As String, atajo As String
    carpeta = RutaWorkspace()
    If Len(carpeta) = 0 Then Exit Sub
    json = LeerArchivo(carpeta & "\settings.json")
    atajo = ValorJsonString(json, "shortcut")
    If Len(atajo) = 0 Then atajo = "Alt+Enter"
    AplicarAtajo atajo
Fin:
End Sub

Private Sub Enviar(ByVal tarea As String, ByVal permiteDetener As Boolean)
    On Error GoTo Fallo
    If Documents.Count = 0 Then Exit Sub

    Dim doc As Document
    Set doc = ActiveDocument

    paso = "detener"
    If HayEnCurso(doc) Then
        If permiteDetener Then
            DetenerPluma
        Else
            Application.StatusBar = "Pluma est" & ChrW(225) & " respondiendo. Presion" & ChrW(225) & " Alt+Enter para detener."
        End If
        Exit Sub
    End If

    Dim ids As String, ajustes As String, carpeta As String
    ids = IdsDeAgentes()
    carpeta = RutaWorkspace()
    If Len(carpeta) > 0 Then ajustes = LeerArchivo(carpeta & "\settings.json")

    Dim instruccion As String, texto As String, linea As String, agente As String
    Dim ancla As Range, p As Paragraph, fuente As Range
    Dim aBorrar As New Collection
    Dim imagenes As Long
    Dim esteAgente As String, resto As String

    paso = "leer seleccion"
    Set fuente = RangoAEnviar(ids).Duplicate
    For Each p In fuente.Paragraphs
        linea = LimpiarLinea(p.Range.Text)
        If SepararAgente(linea, ids, esteAgente, resto) Then
            agente = esteAgente
            If Len(resto) > 0 Then instruccion = resto
            aBorrar.Add p.Range
        Else
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
        Application.StatusBar = "Pluma: no hay texto ni im" & ChrW(225) & "genes para enviar."
        LogMacro "nada para enviar (seleccion tipo " & Selection.Type & ")"
        Exit Sub
    End If
    If Len(texto) > MAX_CHARS Then texto = Left$(texto, MAX_CHARS)
    If Len(agente) = 0 Then
        Dim pref As String
        pref = ValorJsonString(ajustes, "defaultAgent")
        If Len(pref) > 0 Then agente = LCase$(pref)
    End If
    Dim tinta As String, web As String
    tinta = ValorJsonNumero(ajustes, "streamColor")
    web = ValorJsonBool(ajustes, "allowInternet")
    If Len(web) = 0 Then web = "true"

    Dim i As Long
    paso = "borrar mencion"
    For i = aBorrar.Count To 1 Step -1
        aBorrar(i).Delete
    Next i

    Dim id As String
    Randomize
    id = Format$(Now, "yyyymmddhhnnss") & CStr(Int(Rnd * 100000))

    paso = "marcador"
    If imagenes > 0 Then doc.Bookmarks.Add "_pluma_" & id, fuente

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
    cc.Title = "Pluma"
    cc.Range.Text = ChrW(10035) & " Pluma est" & ChrW(225) & " pensando" & ChrW(8230)
    cc.Appearance = WD_CC_HIDDEN
    cc.Range.Font.Color = WD_COLOR_STATUS
    cc.LockContents = True

    paso = "variables"
    If Len(texto) > 0 Then doc.Variables.Add TAG_PREFIX & id & "-text", texto
    If Len(instruccion) > 0 Then doc.Variables.Add TAG_PREFIX & id & "-instr", instruccion
    If Len(agente) > 0 Then doc.Variables.Add TAG_PREFIX & id & "-agent", agente
    If Len(tarea) > 0 And Len(instruccion) = 0 Then doc.Variables.Add TAG_PREFIX & id & "-task", tarea
    If Len(tinta) > 0 Then doc.Variables.Add TAG_PREFIX & id & "-ink", tinta
    doc.Variables.Add TAG_PREFIX & id & "-web", web
    doc.Variables.Add TAG_PREFIX & id & "-color", CStr(ColorDelTexto(ancla))

    Dim exe As String
    exe = Environ$("LOCALAPPDATA") & "\Pluma\Pluma.exe"
    If Len(Dir$(exe)) = 0 Then
        cc.LockContents = False
        cc.Range.Text = "Pluma no est" & ChrW(225) & " instalado en esta PC."
        cc.Delete False
        Exit Sub
    End If

    paso = "lanzar helper"
    Shell """" & exe & """ " & id, vbHide
    Application.StatusBar = "Pluma est" & ChrW(225) & " escribiendo" & ChrW(8230)
    Exit Sub

Fallo:
    Application.StatusBar = "Pluma: " & Err.Description
    LogMacro "error en '" & paso & "': " & Err.Number & " " & Err.Description
End Sub

Private Sub LogMacro(ByVal msg As String)
    On Error Resume Next
    Dim dir_ As String, f As Integer
    dir_ = Environ$("LOCALAPPDATA") & "\Pluma\logs"
    If Len(Dir$(dir_, vbDirectory)) = 0 Then Exit Sub
    f = FreeFile
    Open dir_ & "\macro.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & msg
    Close #f
End Sub

Public Sub DetenerPluma()
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
        Application.StatusBar = "Pluma: no hay ninguna respuesta en curso."
        Exit Sub
    End If
    For Each cc In doc.ContentControls
        If Left$(cc.Tag, Len(TAG_PREFIX)) = TAG_PREFIX Then
            If enCursor = "" Or cc.Tag = enCursor Then PedirParada doc, cc.Tag
        End If
    Next cc
    Application.StatusBar = "Pluma: deteniendo" & ChrW(8230)
    Exit Sub

Fallo:
    Application.StatusBar = "Pluma: " & Err.Description
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

Private Sub PedirParada(doc As Document, ByVal tag As String)
    If HelperVivo(doc, tag) Then
        On Error Resume Next
        doc.Variables.Add tag & "-stop", "1"
    Else
        CerrarHuerfano doc, tag
    End If
End Sub

Private Function HelperVivo(doc As Document, ByVal tag As String) As Boolean
    On Error GoTo NoVivo
    Dim hb As Double
    hb = Val(doc.Variables(tag & "-hb").Value)
    HelperVivo = (CDbl(Now) - hb) * 86400# < 6
    Exit Function
NoVivo:
    HelperVivo = False
End Function

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
    For Each sufijo In Array("-text", "-instr", "-color", "-stop", "-hb", "-agent", "-task", "-ink", "-web")
        doc.Variables(tag & sufijo).Delete
    Next sufijo
    doc.Bookmarks("_pluma_" & Mid$(tag, Len(TAG_PREFIX) + 1)).Delete
End Sub

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

Private Function ColorDelTexto(ByVal r As Range) As Long
    Dim color As Long
    color = PrimerCaracterDeTexto(r).Font.Color
    If color = wdUndefined Then color = wdColorAutomatic
    ColorDelTexto = color
End Function

Private Function RangoAEnviar(ByVal ids As String) As Range
    If Selection.Type <> wdSelectionIP Then
        Set RangoAEnviar = Selection.Range
        Exit Function
    End If

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
    If SoloInstrucciones(r, ids) Then
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

Private Function EsRespuesta(ByVal q As Paragraph) As Boolean
    Dim b As Bookmark
    For Each b In q.Range.Bookmarks
        If Left$(b.Name, Len(RESP_PREFIX)) = RESP_PREFIX Then
            EsRespuesta = True
            Exit Function
        End If
    Next b
End Function

Private Function SoloInstrucciones(ByVal r As Range, ByVal ids As String) As Boolean
    Dim q As Paragraph, a As String, resto As String
    For Each q In r.Paragraphs
        If Not EsVacio(q) Then
            If Not SepararAgente(LimpiarLinea(q.Range.Text), ids, a, resto) Then Exit Function
        End If
    Next q
    SoloInstrucciones = True
End Function

' Una linea es mencion si arranca con @ y el token esta en agents.json.
' agente vuelve en minuscula; resto es la instruccion, sin los dos puntos iniciales.
Private Function SepararAgente(ByVal linea As String, ByVal ids As String, ByRef agente As String, ByRef resto As String) As Boolean
    SepararAgente = False
    agente = ""
    resto = ""
    If Left$(linea, 1) <> "@" Then Exit Function
    Dim i As Long, token As String, c As String
    i = 2
    Do While i <= Len(linea)
        c = Mid$(linea, i, 1)
        If c = " " Or c = ":" Or c = vbTab Then Exit Do
        If Not (c Like "[A-Za-z0-9_-]") Then Exit Function
        token = token & c
        i = i + 1
    Loop
    If Len(token) = 0 Then Exit Function
    If InStr(1, ids, "|" & LCase$(token) & "|", vbTextCompare) = 0 Then Exit Function
    agente = LCase$(token)
    resto = Trim$(Mid$(linea, i))
    If Left$(resto, 1) = ":" Then resto = Trim$(Mid$(resto, 2))
    SepararAgente = True
End Function

Private Function IdsDeAgentes() As String
    On Error GoTo Fallback
    Dim ruta As String, txt As String, ids As String, token As String
    Dim re As Object, ms As Object, i As Long
    ruta = RutaWorkspace()
    If Len(ruta) = 0 Then GoTo Fallback
    txt = LeerArchivo(ruta & "\agents.json")
    If Len(txt) = 0 Then GoTo Fallback
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = """id""\s*:\s*""([A-Za-z0-9_\-]+)"""
    Set ms = re.Execute(txt)
    ids = "|"
    For i = 0 To ms.Count - 1
        token = LCase$(ms(i).SubMatches(0))
        If InStr(1, ids, "|" & token & "|", vbTextCompare) = 0 Then ids = ids & token & "|"
    Next i
    If ids <> "|" Then
        IdsDeAgentes = ids
        Exit Function
    End If
Fallback:
    IdsDeAgentes = "|claude|codex|gemini|grok|hermes|aider|"
End Function

Private Function RutaWorkspace() As String
    On Error GoTo Fin
    Dim txt As String, re As Object, ms As Object
    txt = LeerArchivo(Environ$("LOCALAPPDATA") & "\Pluma\config.json")
    If Len(txt) = 0 Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """workspace""\s*:\s*""((?:\\.|[^""\\])*)"""
    Set ms = re.Execute(txt)
    If ms.Count > 0 Then RutaWorkspace = UnescapeJson(CStr(ms(0).SubMatches(0)))
Fin:
End Function

Private Function UnescapeJson(ByVal s As String) As String
    Dim i As Long, c As String, n As String, out As String
    i = 1
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If c = "\" And i < Len(s) Then
            n = Mid$(s, i + 1, 1)
            If n = "\" Or n = "/" Or n = """" Then
                out = out & n
            ElseIf n = "n" Then
                out = out & vbLf
            ElseIf n = "r" Then
                out = out & vbCr
            ElseIf n = "t" Then
                out = out & vbTab
            Else
                out = out & n
            End If
            i = i + 2
        Else
            out = out & c
            i = i + 1
        End If
    Loop
    UnescapeJson = out
End Function

Private Function LeerArchivo(ByVal ruta As String) As String
    On Error GoTo Fin
    Dim n As Long, f As Integer
    n = FileLen(ruta)
    If n <= 0 Or n > 2000000 Then Exit Function
    f = FreeFile
    Open ruta For Binary Access Read As #f
    LeerArchivo = Space$(n)
    Get #f, , LeerArchivo
    Close #f
Fin:
End Function

Private Sub AplicarAtajo(ByVal atajo As String)
    On Error GoTo Fin
    Dim t As Template
    Set t = PlantillaPluma()
    If t Is Nothing Then Exit Sub
    Dim anterior As Object
    Set anterior = CustomizationContext
    CustomizationContext = t
    Dim i As Long
    For i = KeyBindings.Count To 1 Step -1
        If StrComp(KeyBindings(i).Command, "EnviarAPluma", vbTextCompare) = 0 Then KeyBindings(i).Clear
    Next i
    Dim code As Long
    code = CodigoAtajo(atajo)
    If code <> 0 Then KeyBindings.Add 1, "EnviarAPluma", code
    On Error Resume Next
    t.Save
    CustomizationContext = anterior
Fin:
End Sub

Private Function PlantillaPluma() As Template
    Dim t As Template
    For Each t In Application.Templates
        If StrComp(t.Name, "Pluma.dotm", vbTextCompare) = 0 Then
            Set PlantillaPluma = t
            Exit Function
        End If
    Next t
End Function

' "Ctrl+Alt+Enter" -> BuildKeyCode. Alt=1024, Ctrl=512, Shift=256, Enter=13.
Private Function CodigoAtajo(ByVal s As String) As Long
    Dim partes() As String, i As Long, p As String
    Dim alt As Boolean, ctrl As Boolean, shift As Boolean, tecla As Long
    If Len(s) = 0 Then Exit Function
    partes = Split(s, "+")
    For i = 0 To UBound(partes)
        p = LCase$(Trim$(partes(i)))
        Select Case p
            Case "alt": alt = True
            Case "ctrl", "control": ctrl = True
            Case "shift", "mayus": shift = True
            Case "enter", "intro", "return": tecla = 13
            Case "tab": tecla = 9
            Case "espacio", "space": tecla = 32
            Case "esc", "escape": tecla = 27
            Case Else
                If Len(p) = 1 Then
                    tecla = Asc(UCase$(p))
                ElseIf Left$(p, 1) = "f" Then
                    Dim n As Long
                    n = Val(Mid$(p, 2))
                    If n >= 1 And n <= 12 Then tecla = 111 + n
                End If
        End Select
    Next i
    If tecla = 0 Then Exit Function
    If Not alt And Not ctrl And Not shift Then Exit Function
    If alt And ctrl And shift Then
        CodigoAtajo = BuildKeyCode(tecla, 1024, 512, 256)
    ElseIf alt And ctrl Then
        CodigoAtajo = BuildKeyCode(tecla, 1024, 512)
    ElseIf alt And shift Then
        CodigoAtajo = BuildKeyCode(tecla, 1024, 256)
    ElseIf ctrl And shift Then
        CodigoAtajo = BuildKeyCode(tecla, 512, 256)
    ElseIf alt Then
        CodigoAtajo = BuildKeyCode(tecla, 1024)
    ElseIf ctrl Then
        CodigoAtajo = BuildKeyCode(tecla, 512)
    Else
        CodigoAtajo = BuildKeyCode(tecla, 256)
    End If
End Function

Private Function ValorJsonString(ByVal json As String, ByVal clave As String) As String
    On Error GoTo Fin
    If Len(json) = 0 Or Len(clave) = 0 Then Exit Function
    Dim re As Object, ms As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & clave & """" & "\s*:\s*""([^""]*)"""
    Set ms = re.Execute(json)
    If ms.Count > 0 Then ValorJsonString = CStr(ms(0).SubMatches(0))
Fin:
End Function

Private Function ValorJsonBool(ByVal json As String, ByVal clave As String) As String
    On Error GoTo Fin
    If Len(json) = 0 Or Len(clave) = 0 Then Exit Function
    Dim re As Object, ms As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = """" & clave & """" & "\s*:\s*(true|false)"
    Set ms = re.Execute(json)
    If ms.Count > 0 Then ValorJsonBool = LCase$(CStr(ms(0).SubMatches(0)))
Fin:
End Function

Private Function ValorJsonNumero(ByVal json As String, ByVal clave As String) As String
    On Error GoTo Fin
    If Len(json) = 0 Or Len(clave) = 0 Then Exit Function
    Dim re As Object, ms As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & clave & """" & "\s*:\s*(-?\d+)"
    Set ms = re.Execute(json)
    If ms.Count > 0 Then ValorJsonNumero = CStr(ms(0).SubMatches(0))
Fin:
End Function

Private Function LimpiarLinea(ByVal s As String) As String
    s = Replace(s, vbCr, "")
    s = Replace(s, Chr(7), "")
    s = Replace(s, Chr(11), " ")
    LimpiarLinea = Trim$(s)
End Function
