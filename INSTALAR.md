# Instalación en una PC nueva

Tiempo estimado: 10 minutos. No hace falta ser administrador.

## Qué hace falta

- Windows 10 u 11 con **Word de escritorio**: 2016, 2019, 2021 o Microsoft 365. Word en la web no sirve.
- Una cuenta de Claude con plan **Pro, Max, Team o Enterprise**. El plan gratuito no incluye Claude Code.
- La carpeta del workspace de ella (CLAUDE.md, soul.md, files/, skills), que va aparte y **no está en este ZIP**.

## Paso 1 — Instalar Claude Code

Abrir **PowerShell** (menú Inicio → escribir "PowerShell") y pegar:

```powershell
irm https://claude.ai/install.ps1 | iex
```

Después, en la misma ventana:

```powershell
claude
```

Se abre el navegador para **iniciar sesión con la cuenta de ella**. Cuando termine, escribir `/exit` para salir. Esto se hace una sola vez.

> Git for Windows **no** hace falta: el conector no usa comandos de consola.

## Paso 2 — Copiar el workspace

Copiar la carpeta del workspace a su PC, por ejemplo a `Documentos\Claude-Ella`.

Si se copia a `Documentos\Claude-Word`, en el paso 3 no hace falta indicar la ruta: el instalador la usa por defecto.

## Paso 3 — Instalar el conector

1. **Cerrar Word** por completo.
2. Descomprimir `ClaudeWord-instalador.zip` en cualquier lugar, por ejemplo el Escritorio.
3. Abrir la carpeta `ClaudeWord`:
   - Si el workspace está en `Documentos\Claude-Word` → **doble clic en `instalar.cmd`**.
   - Si está en otra ruta → en esa carpeta, clic derecho en un espacio vacío → *Abrir en Terminal*, y ejecutar:
     ```
     .\instalar.cmd -Workspace "C:\Users\ELLA\Documents\Claude-Ella"
     ```
4. Tiene que terminar con **`Listo`** en verde. La línea `Office` al principio muestra su versión de Word.

Después de instalar, la carpeta descomprimida se puede borrar.

## Paso 4 — Probar

Abrir Word, escribir una pregunta, dejar el cursor en ese párrafo y presionar **Alt+Enter**. También se puede hacer clic derecho → **Enviar a Claude**.

Tiene que aparecer "✳ Claude está pensando…" en gris muy claro, y después la respuesta, escribiéndose palabra por palabra.

## Si algo falla

| Síntoma | Solución |
|---|---|
| El instalador dice "No encuentro claude.exe" | Repetir el paso 1 y cerrar y volver a abrir PowerShell |
| Alt+Enter no hace nada, pero el clic derecho sí | Otro programa usa ese atajo (pasó con AMD Adrenalin y Ctrl+Alt+Shift+K). Buscar el atajo en ese programa, o usar el clic derecho |
| Ni el atajo ni el clic derecho hacen nada | En Word: **Archivo → Opciones → Complementos → Administrar: Plantillas → Ir**. `ClaudeWord.dotm` tiene que estar tildado |
| Aparece "⚠ Claude no pudo responder: … login …" | Correr `claude` en PowerShell y volver a iniciar sesión |
| Cualquier otra cosa | Revisar `%LOCALAPPDATA%\ClaudeWord\logs\last.log` |

## Ajustes sin reinstalar

En `%LOCALAPPDATA%\ClaudeWord\config.json` se cambian la carpeta del workspace, la velocidad de escritura y las pausas. Los cambios valen desde el siguiente pedido. El detalle está en el README.

## Desinstalar

Con Word cerrado, doble clic en `desinstalar.cmd`. El workspace no se toca.
