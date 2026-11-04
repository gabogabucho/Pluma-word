# Pluma

Pluma mete un agente de escritura dentro de Microsoft Word. No usa panel, popups ni terminal a la vista. El agente es el que tengas instalado: Claude Code queda como opción por defecto, y también podés usar Codex CLI, Gemini CLI, Grok CLI, Hermes o aider.

## Uso

1. Dejá el cursor en el mensaje, o en la línea vacía de abajo. Un **mensaje** son los renglones seguidos con texto o imágenes: una línea en blanco separa un mensaje del siguiente. Si seleccionás, se envía exactamente la selección.
2. Presioná **Alt+Enter**, o hacé clic derecho y elegí **Pluma → Enviar**.
3. La respuesta aparece **debajo**, en gris y en stream. Cuando termina, queda como texto normal.

Para elegir agente e instrucción, escribí una línea dentro de la selección:

```
Me gustaría que el informe llegue antes del viernes porque necesitamos revisarlo.
@claude hacelo más formal
```

También valen `@codex`, `@gemini`, `@grok`, `@hermes` y `@aider` (y cualquier otro `id` que agregues en `agents.json`). La línea de la mención se borra sola. Si no hay mención, se usa el agente por defecto y la tarea por defecto del workspace.

**Tareas del menú contextual** (clic derecho → Pluma): Corregir, Más formal, Resumir, Expandir y Traducir. El texto de cada una está en `tasks.json` y se puede editar sin reinstalar. Si en el mismo mensaje hay una instrucción escrita (`@claude hacelo más corto`), esa instrucción manda y la tarea del menú no se aplica.

**Imágenes:** las imágenes pegadas en línea (Ctrl+V de una captura, por ejemplo) viajan cuando el agente las acepta.

| Agente | Imágenes |
|---|---|
| `@claude` | Sí, dentro del mensaje |
| `@codex` | Sí, con `-i` |
| `@hermes` | Una sola, con `--image` |
| `@grok`, `@gemini`, `@aider` | No. Pluma avisa y no envía el pedido |

Las imágenes flotantes, con ajuste de texto, no se envían.

**Formato:** la respuesta copia el estilo, la fuente, el tamaño, el color y la alineación del texto. Si escribiste un título o un ítem de lista, la respuesta va en estilo Normal.

Para **detener** una respuesta, volvé a presionar **Alt+Enter** en cualquier lugar del documento, o clic derecho → Pluma → Detener. Mientras hay una respuesta en curso, Alt+Enter detiene en lugar de enviar. Lo ya escrito se conserva y termina en " …". Si el agente todavía estaba pensando, el bloque desaparece. Si Word se cerró a mitad de una respuesta, Detener cierra el bloque que quedó.

## Agentes

El archivo `agents.json` del workspace declara cada proveedor: comando, argumentos, cómo viaja el prompt (`stdin`, `arg` o `file`), si la salida es texto plano o un stream JSON, y el parser.

| Id | Programa | Prompt | Salida | Parser |
|---|---|---|---|---|
| `claude` | Claude Code | stdin, mensaje stream-json (con imágenes) | stream-json | `claude-stream` |
| `codex` | Codex CLI `exec -` | stdin | JSONL | `codex-jsonl` |
| `gemini` | Gemini CLI | stdin | stream-json | `gemini-stream` |
| `grok` | Grok CLI | archivo (`--prompt-file`) | `streaming-json` | `grok-stream` |
| `hermes` | Hermes (`chat --oneshot`) | archivo (`--query-file`) | `stream-json` | `hermes-stream` |
| `aider` | aider | archivo (`--message-file`) | texto, al final | `text` |

Claude Code no cambia de comportamiento: stream, Alt+Enter, detener, formato y las imágenes en el mensaje. Sigue leyendo el `CLAUDE.md` del workspace y no recibe un preámbulo extra. El resto recibe, delante del pedido, las reglas de Word y, si existen, `PLUMA.md` y `context/estilo.md`.

Grok y Hermes tienen modo headless documentado. Los dos están habilitados en el preset, no comentados. Grok no lee el prompt por stdin (hay que pasarlo con `--prompt-file`) y no documenta un flag de imagen; por eso `images` queda en `none`. Hermes acepta una imagen con `--image`. El campo `note` de cada uno, en `agents.json`, resume eso. JSON no admite comentarios `//`.

Codex corre en sandbox de solo lectura. Grok queda limitado a `read_file`, `grep` y `list_dir`. Hermes, con `--yolo`, no se frena a pedir permiso en una consola que no existe: puede usar las herramientas que tenga configuradas, incluidas las que escriben archivos. Si querés achicarlas, sumá `--toolsets` en los `args` de `agents.json`.

Para sumar otro CLI, agregá un objeto en `agents` con `id`, `command`, `args`, `promptVia` y `parser` (`text`, `claude-stream`, `jsonl`, `codex-jsonl`, `gemini-stream`, `grok-stream` o `hermes-stream`). Con `parser: "jsonl"` definís `textRules`, `errorRules` y `statusRules` (`type`, `path`, `mode`, `match`, y `clear` para tirar el texto anterior a una herramienta). En los argumentos, `{run}` es una carpeta temporal del pedido, `{rules}` el archivo de reglas de Word y `{workspace}` la carpeta del workspace.

## Arquitectura

```
Word ── Pluma.dotm (carpeta STARTUP de Word, ubicación de confianza)
          │ macro EnviarAPluma:
          │   · lee la selección, la mención @agente y la tarea del menú
          │   · crea un bloque protegido (content control) debajo
          │   · guarda texto, instrucción, agente y tarea en variables del documento
          ▼
        Pluma.exe <id>   (oculto, un proceso por pedido)
          │ el comando de agents.json, con cwd = workspace
          └─► escribe en el bloque vía COM cada 40 ms y reintenta si Word está ocupado
```

- **El workspace** (por defecto `Documentos\Pluma`) contiene `PLUMA.md` (rol compartido), `CLAUDE.md` (Claude Code lo carga solo e importa `PLUMA.md`), `agents.json`, `tasks.json`, `context/` y `files/`.
- **Claude Code** solo puede leer (`Read`, `Glob`, `Grep`) y usar las skills del workspace (`Skill`). No edita ni ejecuta nada.
- **Formato:** las reglas de salida para Word (texto plano, sin preámbulos) están en `src/helper/word-rules.md`.

### Estructura del workspace

```
Pluma/
├── PLUMA.md                  ← rol para todos los agentes
├── CLAUDE.md                 ← solo lo carga Claude Code; importa PLUMA.md
├── agents.json               ← proveedores
├── tasks.json                ← textos de Corregir, Más formal, etc.
├── context/                  ← notas; Claude las importa, el resto las recibe si están
├── files/                    ← documentos de referencia, a demanda
└── .claude/skills/<nombre>/SKILL.md
```

Claude Code **solo** carga automáticamente los archivos `CLAUDE.md`. Todo otro archivo que tenga que estar siempre presente se importa con `@archivo.md` desde el `CLAUDE.md`.

## Instalación (una sola vez)

Requisitos: Word de escritorio en Windows (2016, 2019, 2021 o 365) y al menos un CLI de la tabla, con la sesión ya iniciada.

```powershell
# Con Word cerrado
powershell -ExecutionPolicy Bypass -File .\install.ps1
# Opcional: -Workspace "C:\ruta\al\workspace" -Model sonnet -DefaultAgent claude -ClaudeConfigDir "C:\ruta"
```

El instalador dice cuáles encontró y deja habilitados solo esos en `agents.json`. Si el archivo ya existía, no lo pisa.

Después editá `PLUMA.md` y `context/estilo.md` en el workspace.

Para desinstalar: `.\uninstall.ps1`. El workspace no se borra. Si había un `ClaudeWord.dotm` de una versión anterior, también se saca para que no se quede con Alt+Enter.

## Configuración

Está en `%LOCALAPPDATA%\Pluma\config.json`:

| Clave | Uso |
|---|---|
| `claudePath` | Si el archivo existe, pisa el `command` de Claude en `agents.json` |
| `workspace` | Carpeta con el rol, las tareas y `agents.json` |
| `model` | Modelo de Claude si el agente no trae uno (`sonnet`, `opus`, etc.) |
| `defaultAgent` | Id que se usa cuando el mensaje no tiene `@agente` |
| `claudeConfigDir` | Opcional. Aísla la configuración de Claude (ver abajo) |
| `timeoutSeconds` | Tiempo máximo por pedido (600 por defecto) |
| `charsPerSecond` | Velocidad de tipeo de la respuesta (20). Más bajo = más lento |
| `startDelayMs` | Pausa en "pensando…" antes de la primera palabra (800) |
| `sentencePauseMs` | Pausa tras `. ? ! :` (450). Tras `, ;` es un tercio |
| `paragraphPauseMs` | Pausa entre párrafos (900) |
| `catchUpAfterChars` | Si el atraso supera esta cantidad de caracteres, acelera en proporción (600) |

El ritmo solo cambia **cómo se muestra** la respuesta. El agente termina a su velocidad y el texto se sigue tipeando. Los cambios en `config.json` y en `agents.json` se aplican en el siguiente pedido, sin reinstalar ni reiniciar Word.

El modelo de Codex, Gemini, Grok, Hermes o aider se pone en el agente (`"model": "..."` junto con `"modelArg": "--model"`), no en `model` de `config.json`. Ese campo es de Claude.

**Aislamiento de Claude:** el helper ya ignora los settings, hooks y servidores MCP del usuario. Pero el `~/.claude/CLAUDE.md` global **sí se carga**. Si en esa PC Claude Code se usa para otras cosas, apuntá `claudeConfigDir` a una carpeta dedicada e iniciá sesión una vez con `set CLAUDE_CONFIG_DIR=<carpeta>` seguido de `claude`.

## Cómo probarlo en Windows

1. Cerrá Word.
2. En PowerShell, desde esta carpeta: `powershell -ExecutionPolicy Bypass -File .\install.ps1`.
3. Tiene que listar los CLI que encontró y terminar con `Listo`.
4. Opcional, sin abrir Word: `%LOCALAPPDATA%\Pluma\Pluma.exe --self-test`. Tiene que imprimir `self-test ok` y devolver código 0. Prueba los parsers (incluido Grok y Hermes), el armado del prompt de Claude y que `agents.json` coincide con el código. No habla con ningún CLI ni con Word.
5. Abrí Word, escribí un párrafo y presioná Alt+Enter. Tiene que aparecer el estado en gris y después la respuesta.
6. En otro párrafo, escribí `@grok resumí esto` (o `@hermes`, `@codex`, `@gemini`) y volvé a enviar. La línea de la mención desaparece.
7. Seleccioná un párrafo, clic derecho → Pluma → Corregir. La respuesta tiene que corregir, no contestar como charla.
8. Mientras escribe, Alt+Enter corta y deja lo ya escrito con " …".
9. Si algo falla, el detalle queda en `%LOCALAPPDATA%\Pluma\logs\last.log` y, si falló la macro, en `macro.log`.

## Diagnóstico

| Síntoma | Qué mirar |
|---|---|
| El instalador no encuentra ningún CLI | Instalá uno (el mensaje del instalador trae el comando) y abrí una terminal nueva |
| Alt+Enter no hace nada, pero el clic derecho sí | Otro programa usa ese atajo. Usá el clic derecho |
| Ni el atajo ni el menú hacen nada | En Word: Archivo → Opciones → Complementos → Administrar: Plantillas → Ir. `Pluma.dotm` tiene que estar tildado |
| "no está habilitado" | Ese CLI no estaba en la PC cuando instalaste. Instalálo y, o bien borrá `agents.json` y reinstalá, o poné `"enabled": true` y la ruta en `"command"` |
| Una imagen con `@grok` o `@gemini` | Esos agentes no reciben imágenes. Usá `@claude`, `@codex` o `@hermes` (una sola) |

Los errores también se escriben dentro del documento con el prefijo ⚠.
