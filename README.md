# Claude para Word

Claude Code local dentro de Microsoft Word. No usa panel, popups ni terminal a la vista.

## Uso

1. Seleccioná un párrafo, o dejá el cursor dentro de él.
2. Presioná **Alt+Enter**, o hacé clic derecho y elegí **Enviar a Claude**.
3. La respuesta aparece **debajo**, en gris y en stream. Cuando termina, queda como texto normal.

Para darle una instrucción, escribila en una línea dentro de la selección:

```
Me gustaría que el informe llegue antes del viernes porque necesitamos revisarlo.
@claude hacelo más formal
```

La línea `@claude` se borra sola. Si no hay línea `@claude`, se aplica la **tarea por defecto** del `CLAUDE.md`.

**Imágenes:** las imágenes pegadas en línea (Ctrl+V de una captura, por ejemplo) viajan junto con el texto. Se puede mandar solo una imagen, o una imagen con `@claude ¿qué dice esto?` debajo. Las imágenes flotantes, con ajuste de texto, no se envían.

**Formato:** la respuesta copia el estilo, la fuente, el tamaño, el color y la alineación del texto de ella. Si ella escribió un título o un ítem de lista, la respuesta va en estilo Normal.

Para **cancelar**, poné el cursor dentro del bloque gris y volvé a presionar el atajo.

## Arquitectura

```
Word ── ClaudeWord.dotm (carpeta STARTUP de Word, ubicación de confianza)
          │ macro EnviarAClaude:
          │   · lee la selección y separa la línea @claude
          │   · crea un bloque protegido (content control) debajo
          │   · guarda texto e instrucción en variables del documento
          ▼
        ClaudeWord.exe <id>   (oculto, un proceso por pedido)
          │ claude -p --output-format stream-json ...  (cwd = workspace)
          └─► escribe en el bloque vía COM cada 150 ms y reintenta si Word está ocupado
```

- **El workspace** (por defecto `Documentos\Claude-Word`) contiene `CLAUDE.md` (rol, soul y tarea por defecto), `context/` y `files/`. Claude Code carga el `CLAUDE.md` automáticamente, y lo que se importa con `@context/...` está presente en cada pedido.
- **Herramientas:** Claude solo puede leer (`Read`, `Glob`, `Grep`) y usar las skills del workspace (`Skill`). No edita ni ejecuta nada.

### Estructura recomendada del workspace

```
Claude-Word/
├── CLAUDE.md                 ← punto de entrada: importa lo que se carga siempre
├── soul.md                   ← importado con @soul.md (no se carga solo)
├── context/                  ← notas; importá con @ lo esencial, el resto a demanda
├── files/                    ← documentos de referencia, se leen a demanda
└── .claude/skills/<nombre>/SKILL.md   ← procedimientos que Claude carga cuando aplican
```

Claude Code **solo** carga automáticamente los archivos `CLAUDE.md`. Todo otro archivo que tenga que estar siempre presente se importa con `@archivo.md` desde el `CLAUDE.md`.
- **Formato:** las reglas de salida para Word (texto plano, sin preámbulos) están en `src/helper/word-rules.md`.

## Instalación (una sola vez)

Requisitos: Word de escritorio en Windows (2016, 2019, 2021 o 365) y Claude Code instalado con la sesión ya iniciada (correr `claude` una vez).

```powershell
# Con Word cerrado
powershell -ExecutionPolicy Bypass -File .\install.ps1
# Opcional: -Workspace "D:\Claude" -Model sonnet -ClaudeConfigDir "C:\ruta"
```

Después editá `CLAUDE.md` y `context/estilo.md` en el workspace.

Para desinstalar: `.\uninstall.ps1`. El workspace no se borra.

## Configuración

Está en `%LOCALAPPDATA%\ClaudeWord\config.json`:

| Clave | Uso |
|---|---|
| `claudePath` | Ruta a `claude.exe` |
| `workspace` | Carpeta con el rol, el soul y el contexto |
| `model` | Modelo opcional (`sonnet`, `opus`, etc.) |
| `claudeConfigDir` | Opcional. Aísla la configuración de Claude (ver abajo) |
| `timeoutSeconds` | Tiempo máximo por pedido (600 por defecto) |
| `charsPerSecond` | Velocidad de tipeo de la respuesta (20). Más bajo = más lento |
| `startDelayMs` | Pausa en "pensando…" antes de la primera palabra (800) |
| `sentencePauseMs` | Pausa tras `. ? ! :` (450). Tras `, ;` es un tercio |
| `paragraphPauseMs` | Pausa entre párrafos (900) |
| `catchUpAfterChars` | Si el atraso supera esta cantidad de caracteres, acelera en proporción (600) |

El ritmo solo cambia **cómo se muestra** la respuesta, no cuánto consume: Claude termina a su velocidad y el texto se sigue tipeando. Los cambios en `config.json` se aplican en el siguiente pedido, sin reinstalar ni reiniciar Word.

**Aislamiento:** el helper ya ignora los settings, hooks y servidores MCP del usuario. Pero el `~/.claude/CLAUDE.md` global **sí se carga**. En una PC donde Claude Code se usa para otras cosas, apuntá `claudeConfigDir` a una carpeta dedicada e iniciá sesión una vez con `set CLAUDE_CONFIG_DIR=<carpeta>` seguido de `claude`.

## Diagnóstico

Si algo falla, el log del último pedido está en `%LOCALAPPDATA%\ClaudeWord\logs\last.log`. Los errores también se escriben dentro del documento con el prefijo ⚠.
