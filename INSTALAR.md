# Instalación en una PC nueva

Unos minutos. No hace falta ser administrador.

## Qué hace falta

- Windows 10 u 11 con **Word de escritorio**: 2016, 2019, 2021 o Microsoft 365. Word en la web no sirve.
- Al menos uno de estos programas, con la sesión ya iniciada:
  - **Claude Code** (hace falta un plan Pro, Max, Team o Enterprise; el plan gratuito no incluye Claude Code)
  - **Codex CLI**
  - **Gemini CLI**
  - **Grok CLI**
  - **Hermes**, de Nous Research
  - **aider**
- Si ya tenés un workspace armado (rol, notas, archivos), copialo aparte. No viene en el ZIP.

## Paso 1 — Instalar un agente

Abrí **PowerShell** (menú Inicio → escribir "PowerShell") y pegá el que vayas a usar. Podés instalar más de uno: el instalador de Pluma detecta cuáles están.

```powershell
# Claude Code
irm https://claude.ai/install.ps1 | iex
claude

# Codex CLI
npm install -g @openai/codex

# Gemini CLI
npm install -g @google/gemini-cli

# Grok CLI
irm https://x.ai/cli/install.ps1 | iex
grok

# Hermes
iex (irm https://hermes-agent.nousresearch.com/install.ps1)

# aider
pipx install aider-chat
```

Con Claude y con Grok, la primera vez se abre el navegador para iniciar sesión. Cuando termine, salí (`/exit` en Claude). Con Hermes, seguí lo que pida el instalador para elegir el modelo. Esto se hace una sola vez.

Cerrá esa ventana de PowerShell y abrí otra, así toma el PATH nuevo.

> Git for Windows no hace falta para Pluma. Algún instalador de los CLI puede pedirlo por su cuenta.

## Paso 2 — El workspace

Si ya tenés la carpeta con el rol y las notas, copiala a la PC, por ejemplo a `Documentos\Pluma`.

Si no existe, el instalador crea una plantilla en `Documentos\Pluma`. Después editá `PLUMA.md` y `context/estilo.md`.

## Paso 3 — Instalar Pluma

1. **Cerrá Word** por completo.
2. Descomprimí `Pluma-instalador.zip` en cualquier lugar, por ejemplo el Escritorio.
3. Abrí la carpeta `Pluma`:
   - Si el workspace está en `Documentos\Pluma` → **doble clic en `instalar.cmd`**.
   - Si está en otra ruta → en esa carpeta, clic derecho en un espacio vacío → *Abrir en Terminal*, y ejecutá:
     ```
     .\instalar.cmd -Workspace "C:\ruta\al\workspace"
     ```
4. Al principio lista cada CLI: la ruta si lo encontró, o `(no instalado)`. Tiene que terminar con **`Listo`** en verde. La línea `Office` muestra la versión de Word.

El agente por defecto es el primero que encuentre, en este orden: Claude, Codex, Gemini, Grok, Hermes, aider. Para forzar uno: `.\instalar.cmd -DefaultAgent grok`.

Después de instalar, la carpeta descomprimida se puede borrar. El workspace no.

Si en esa PC había una versión anterior (Claude para Word), el instalador saca `ClaudeWord.dotm` para que Alt+Enter no quede enganchado dos veces. La carpeta vieja `%LOCALAPPDATA%\ClaudeWord` no se borra sola: la podés eliminar a mano si ya no la usás.

## Paso 4 — Probar

Abrí Word, escribí una pregunta, dejá el cursor en ese párrafo y presioná **Alt+Enter**. También podés hacer clic derecho → **Pluma** → **Enviar**.

Tiene que aparecer un aviso en gris muy claro y, después, la respuesta escribiéndose.

Para otro agente, en el mismo párrafo:

```
@grok resumí esto en tres oraciones
```

La línea `@grok` desaparece y responde Grok. Lo mismo con `@hermes`, `@codex`, `@gemini`, `@claude` y `@aider`.

Para una tarea fija, seleccioná el texto y usá clic derecho → Pluma → Corregir, Más formal, Resumir, Expandir o Traducir.

## Si algo falla

| Síntoma | Solución |
|---|---|
| El instalador dice que no hay ningún CLI | Repetí el paso 1 y abrí una PowerShell nueva |
| Alt+Enter no hace nada, pero el clic derecho sí | Otro programa usa ese atajo. Usá el clic derecho |
| Ni el atajo ni el clic derecho hacen nada | En Word: **Archivo → Opciones → Complementos → Administrar: Plantillas → Ir**. `Pluma.dotm` tiene que estar tildado |
| "no está habilitado" o "no encontré" | Ese programa no estaba instalado. Instalálo y volvé a correr el instalador después de borrar `agents.json`, o editá la ruta y poné `"enabled": true` |
| Aparece un aviso de login o de sesión | Abrí el CLI en PowerShell (`claude`, `grok`, `hermes`, etc.) e iniciá sesión de nuevo |
| Cualquier otra cosa | Revisá `%LOCALAPPDATA%\Pluma\logs\last.log` |

## Ajustes sin reinstalar

En `%LOCALAPPDATA%\Pluma\config.json` se cambian la carpeta del workspace, el agente por defecto y la velocidad de escritura. En el `agents.json` del workspace se cambian los comandos. Los dos valen desde el siguiente pedido. El detalle está en el README.

## Desinstalar

Con Word cerrado, doble clic en `desinstalar.cmd`. El workspace no se toca.
