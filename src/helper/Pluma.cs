// Pluma.exe — helper oculto que ejecuta un agente CLI y escribe la
// respuesta en stream dentro de Word via COM.
//
// Lo lanza la macro EnviarAPluma con un id. Con ese id encuentra el bloque
// (content control con tag "pluma-<id>") y las variables del documento.
// El agente (comando, argumentos, como pasa el prompt y el parser) sale de
// agents.json en el workspace. Sin ese archivo, queda el comportamiento
// historico de Claude Code.
//
// Compila con el csc.exe de .NET Framework 4 (C# 5: sin interpolacion,
// sin ?., sin nameof).

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using System.Xml;

namespace Pluma
{
    static class Program
    {
        const string TagPrefix = "pluma-";
        const int TickMs = 40;
        const int ColorStatus = 15263976;       // #E8E8E8: avisos de estado, se ven apenas
        const int ColorAutomatic = -16777216;   // wdColorAutomatic
        const string StatusReadingClaude = "✳ Claude está leyendo el contexto…";
        const int MaxArgPrompt = 7000;

        static readonly object Gate = new object();
        static readonly StringBuilder Answer = new StringBuilder();
        static readonly StringBuilder StdErr = new StringBuilder();
        static int answerColor = ColorAutomatic;
        static string status;
        static string resultText;
        static string error;
        static bool finished;
        static string logPath;

        [STAThread]
        static int Main(string[] args)
        {
            if (args.Length >= 1 && args[0] == "--self-test") return SelfTest();

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            Directory.CreateDirectory(Path.Combine(baseDir, "logs"));
            logPath = Path.Combine(baseDir, "logs", "last.log");
            File.WriteAllText(logPath, DateTime.Now + " inicio\r\n");

            if (args.Length < 1) { Log("falta el id"); return 2; }
            string tag = TagPrefix + args[0];

            IOleMessageFilter oldFilter;
            CoRegisterMessageFilter(new RetryFilter(), out oldFilter);

            dynamic doc = null;
            try
            {
                dynamic word = Marshal.GetActiveObject("Word.Application");
                doc = FindDocument(word, tag);
                if (doc == null) { Log("no encontre el bloque " + tag); return 3; }
                return Run(doc, tag, baseDir);
            }
            catch (Exception ex)
            {
                Log("error: " + ex);
                if (doc != null) Finish(doc, tag, "⚠ Pluma no pudo responder: " + ex.Message, ColorAutomatic);
                return 1;
            }
        }

        static int Run(dynamic doc, string tag, string baseDir)
        {
            Config cfg = Config.Load(Path.Combine(baseDir, "config.json"));
            string texto = ReadVariable(doc, tag + "-text") ?? "";
            string instr = ReadVariable(doc, tag + "-instr");
            string agentId = ReadVariable(doc, tag + "-agent");
            string taskId = ReadVariable(doc, tag + "-task");
            int color;
            if (int.TryParse((string)ReadVariable(doc, tag + "-color"), out color)) answerColor = color;

            AgentFile agents;
            try { agents = AgentFile.Load(Path.Combine(cfg.Workspace ?? "", "agents.json"), cfg); }
            catch (Exception ex)
            {
                Finish(doc, tag, "⚠ Pluma no pudo leer agents.json: " + ex.Message, answerColor);
                return 4;
            }

            if (string.IsNullOrEmpty(agentId) && !string.IsNullOrEmpty(cfg.DefaultAgent))
                agentId = cfg.DefaultAgent;
            Agent agent = agents.Resolve(agentId);
            if (agent == null)
            {
                Finish(doc, tag, "⚠ No existe el agente \"" + agentId + "\". En agents.json están: " + agents.Ids() + ".", answerColor);
                return 4;
            }
            if (!agent.Enabled)
            {
                Finish(doc, tag, "⚠ " + agent.Label + " no está habilitado. El instalador no lo encontró; revisá agents.json.", answerColor);
                return 4;
            }
            if (!KnownParser(agent.Parser))
            {
                Finish(doc, tag, "⚠ " + agent.Label + " tiene un parser desconocido (" + agent.Parser + ").", answerColor);
                return 4;
            }

            if (string.IsNullOrEmpty(instr) && !string.IsNullOrEmpty(taskId))
            {
                instr = LookupTask(cfg.Workspace, taskId);
                if (instr == null)
                {
                    Finish(doc, tag, "⚠ No existe la tarea \"" + taskId + "\". Revisá tasks.json.", answerColor);
                    return 4;
                }
            }

            string command = ResolveCommand(agent, cfg);
            if (command.Length == 0 || (IsRootedCommand(command) && !File.Exists(command)))
            {
                if (agent.Id == "claude")
                    Finish(doc, tag, "⚠ No encontré Claude Code en esta PC (" + command + ").", answerColor);
                else
                    Finish(doc, tag, "⚠ No encontré " + agent.Label + " en esta PC (" + command + ").", answerColor);
                return 4;
            }

            if (string.IsNullOrEmpty(cfg.Workspace) || !Directory.Exists(cfg.Workspace))
            {
                Finish(doc, tag, "⚠ No existe la carpeta del workspace (" + cfg.Workspace + ").", answerColor);
                return 4;
            }

            List<Attachment> images = Attachment.FromBookmark(doc, BookmarkName(tag));
            Log("agente=" + agent.Id + " imagenes=" + images.Count);
            if (images.Count == 0 && texto.Trim().Length == 0)
            {
                Finish(doc, tag, "⚠ No pude leer la imagen. Probá pegarla de nuevo (Ctrl+V) y volvé a enviarla.", answerColor);
                return 5;
            }
            if (images.Count > 0 && agent.Images != "claude" && agent.Images != "files")
            {
                Finish(doc, tag, "⚠ " + agent.Label + " no recibe imágenes. Sacá la imagen o usá un agente que sí (@claude, @codex o @hermes).", answerColor);
                return 5;
            }
            if (agent.MaxImages > 0 && images.Count > agent.MaxImages)
            {
                string n = agent.MaxImages == 1 ? "una imagen" : agent.MaxImages + " imágenes";
                Finish(doc, tag, "⚠ " + agent.Label + " recibe como máximo " + n + " por pedido.", answerColor);
                return 5;
            }

            string runDir = null;
            Process p = null;
            try
            {
                if (NeedsRunDir(agent, images.Count))
                {
                    runDir = Path.Combine(Path.GetTempPath(), "pluma-" + Guid.NewGuid().ToString("n"));
                    Directory.CreateDirectory(runDir);
                }

                var imagePaths = new List<string>();
                if (agent.Images == "files")
                {
                    for (int i = 0; i < images.Count; i++)
                    {
                        string path = Path.Combine(runDir, "img" + (i + 1) + ExtensionFor(images[i].MediaType));
                        File.WriteAllBytes(path, Convert.FromBase64String(images[i].Base64));
                        imagePaths.Add(path);
                    }
                }

                string textPrompt = ComposePrompt(agent, cfg, baseDir, instr, texto, images.Count);
                string problem = PromptProblem(agent, textPrompt);
                if (problem != null)
                {
                    Finish(doc, tag, "⚠ " + agent.Label + " " + problem, answerColor);
                    return 4;
                }

                string stdinPayload = null;
                if (agent.PromptVia == "stdin")
                {
                    stdinPayload = agent.PromptFormat == "claude-message"
                        ? BuildClaudeMessage(textPrompt, images)
                        : textPrompt;
                }
                else if (agent.PromptVia == "file")
                {
                    File.WriteAllText(Path.Combine(runDir, "prompt.txt"), textPrompt, new UTF8Encoding(false));
                }

                Launch launch = BuildLaunch(agent, cfg, baseDir, command, textPrompt, imagePaths, runDir);
                ResetRun();
                lock (Gate) { status = Thinking(agent); }

                p = StartProcess(launch, stdinPayload, agent);
                Stopwatch clock = Stopwatch.StartNew();
                var typer = new Typewriter(cfg);
                string shown = null;
                bool showingStatus = false;
                long firstTextAt = -1;
                long lastBeat = -1000;
                string target = "";

                while (true)
                {
                    bool done; string text, st;
                    lock (Gate)
                    {
                        done = finished;
                        text = Answer.Length > 0 ? Answer.ToString() : (done ? (resultText ?? "") : "");
                        st = status;
                    }

                    dynamic cc = GetBlock(doc, tag);
                    if (cc == null)
                    {
                        Log("bloque borrado: se cancela");
                        Kill(p);
                        DeleteVariables(doc, tag);
                        return 0;
                    }

                    if (clock.ElapsedMilliseconds - lastBeat >= 1000)
                    {
                        SetVariable(doc, tag + "-hb", DateTime.Now.ToOADate().ToString(CultureInfo.InvariantCulture));
                        lastBeat = clock.ElapsedMilliseconds;
                    }

                    if (ReadVariable(doc, tag + "-stop") != null)
                    {
                        Log("detenido");
                        Kill(p);
                        if (shown == null || showingStatus) FinishEmpty(doc, tag);
                        else Finish(doc, tag, shown + " …", answerColor);
                        return 0;
                    }

                    target = Clean(text);
                    if (target.Length == 0)
                    {
                        if (st != null && st != shown)
                        {
                            WriteBlock(cc, st, ColorStatus);
                            shown = st;
                            showingStatus = true;
                            typer.Reset();
                            firstTextAt = -1;
                        }
                        if (done) break;
                    }
                    else
                    {
                        if (firstTextAt < 0) firstTextAt = clock.ElapsedMilliseconds;
                        int len = clock.ElapsedMilliseconds - firstTextAt >= cfg.StartDelayMs
                            ? typer.Next(target, done) : 0;
                        if (len > 0)
                        {
                            string visible = target.Substring(0, len);
                            if (visible != shown)
                            {
                                if (shown != null && !showingStatus && visible.StartsWith(shown, StringComparison.Ordinal))
                                    AppendBlock(cc, visible.Substring(shown.Length));
                                else
                                    WriteBlock(cc, visible, answerColor);
                                shown = visible;
                                showingStatus = false;
                            }
                        }
                        if (done && len >= target.Length) break;
                    }

                    if (clock.Elapsed.TotalSeconds > cfg.TimeoutSeconds)
                    {
                        Kill(p);
                        lock (Gate) { error = "tardó más de " + cfg.TimeoutSeconds + " segundos"; }
                        break;
                    }
                    Thread.Sleep(TickMs);
                }

                if (p != null) p.WaitForExit(5000);

                string err;
                lock (Gate)
                {
                    err = error;
                    if (err == null && target.Length == 0 && StdErr.Length > 0) err = StdErr.ToString().Trim();
                    if (err == null && target.Length == 0 && p != null && p.HasExited && p.ExitCode != 0)
                        err = "terminó con código " + p.ExitCode;
                }
                Log("fin. error=" + (err ?? "-") + " chars=" + target.Length);

                if (err != null && target.Length == 0)
                    Finish(doc, tag, "⚠ " + agent.Label + " no pudo responder: " + err, answerColor);
                else
                    Finish(doc, tag, target == shown ? null : target, answerColor);
                return 0;
            }
            finally
            {
                if (runDir != null)
                {
                    try { Directory.Delete(runDir, true); } catch { }
                }
            }
        }

        // ---------- agentes ----------

        static string[] ClaudeArgs()
        {
            return new[] {
                "-p",
                "--input-format", "stream-json",
                "--output-format", "stream-json",
                "--include-partial-messages",
                "--verbose",
                "--setting-sources", "project",
                "--strict-mcp-config",
                "--no-session-persistence",
                "--tools", "Read,Glob,Grep,Skill"
            };
        }

        static bool KnownParser(string parser)
        {
            return parser == "text" || parser == "claude-stream" || parser == "jsonl"
                || parser == "codex-jsonl" || parser == "gemini-stream"
                || parser == "grok-stream" || parser == "hermes-stream";
        }

        static bool NeedsRunDir(Agent agent, int imageCount)
        {
            if (agent.PromptVia == "file") return true;
            if (agent.Images == "files" && imageCount > 0) return true;
            foreach (string a in agent.Args)
            {
                if (a != null && a.IndexOf("{run}", StringComparison.Ordinal) >= 0) return true;
            }
            return false;
        }

        static string ExtensionFor(string mediaType)
        {
            if (mediaType == "image/jpeg" || mediaType == "image/jpg") return ".jpg";
            if (mediaType == "image/gif") return ".gif";
            return ".png";
        }

        static string ResolveCommand(Agent agent, Config cfg)
        {
            if (agent.Id == "claude" && !string.IsNullOrEmpty(cfg.ClaudePath) && File.Exists(cfg.ClaudePath))
                return cfg.ClaudePath;
            return agent.Command ?? "";
        }

        static string ResolveModel(Agent agent, Config cfg)
        {
            if (!string.IsNullOrEmpty(agent.Model)) return agent.Model;
            if (agent.Id == "claude" && !string.IsNullOrEmpty(cfg.Model)) return cfg.Model;
            return "";
        }

        static bool IsRootedCommand(string command)
        {
            if (string.IsNullOrEmpty(command)) return false;
            if (command.IndexOf('\\') >= 0 || command.IndexOf('/') >= 0) return true;
            return Path.IsPathRooted(command);
        }

        static bool IsCmdScript(string command)
        {
            string ext = Path.GetExtension(command ?? "").ToLowerInvariant();
            return ext == ".cmd" || ext == ".bat";
        }

        static string PromptProblem(Agent agent, string prompt)
        {
            if (agent.PromptVia != "stdin" && agent.PromptVia != "arg" && agent.PromptVia != "file")
                return "tiene promptVia desconocido (" + agent.PromptVia + "). Usá stdin, arg o file.";
            if (agent.PromptVia == "arg" && prompt != null && prompt.Length > MaxArgPrompt)
                return "recibe el prompt como argumento y este texto supera los " + MaxArgPrompt
                    + " caracteres. En agents.json poné \"promptVia\": \"stdin\" o \"file\".";
            if (agent.PromptVia == "file" && string.IsNullOrEmpty(agent.PromptArg))
                return "recibe el prompt en un archivo pero le falta promptArg (por ejemplo --message-file).";
            if (agent.Images == "files" && string.IsNullOrEmpty(agent.ImageArg))
                return "declara imágenes por archivo pero le falta imageArg (por ejemplo -i).";
            return null;
        }

        static string Subst(string s, Config cfg, string runDir, string rulesPath)
        {
            if (s == null) return "";
            return s.Replace("{run}", runDir ?? "")
                    .Replace("{rules}", rulesPath ?? "")
                    .Replace("{workspace}", cfg.Workspace ?? "")
                    .Replace("{claudeConfigDir}", cfg.ClaudeConfigDir ?? "");
        }

        static Dictionary<string, string> ResolveEnv(Agent agent, Config cfg, string runDir, string rulesPath)
        {
            var env = new Dictionary<string, string>();
            foreach (KeyValuePair<string, string> kv in agent.Env)
            {
                string val = Subst(kv.Value, cfg, runDir, rulesPath);
                if (val.Length > 0) env[kv.Key] = val;
            }
            if (agent.Id == "claude" && !string.IsNullOrEmpty(cfg.ClaudeConfigDir) && !env.ContainsKey("CLAUDE_CONFIG_DIR"))
                env["CLAUDE_CONFIG_DIR"] = cfg.ClaudeConfigDir;
            return env;
        }

        // Texto que ve el modelo. Claude Code no lleva preámbulo: el CLAUDE.md
        // del workspace ya entra solo. El resto recibe las reglas de Word y,
        // si existen, PLUMA.md y context/estilo.md.
        static string ComposePrompt(Agent agent, Config cfg, string baseDir, string instr, string texto, int imageCount)
        {
            string user = BuildPrompt(instr, texto, imageCount);
            if (agent.Rules != "prepend" && agent.Context != "prepend") return user;
            var sb = new StringBuilder();
            if (agent.Rules == "prepend")
            {
                string rulesPath = Path.Combine(baseDir, "word-rules.md");
                if (File.Exists(rulesPath)) sb.Append(File.ReadAllText(rulesPath, Encoding.UTF8).Trim()).Append("\n\n");
            }
            if (agent.Context == "prepend" && !string.IsNullOrEmpty(cfg.Workspace))
            {
                AppendFile(sb, Path.Combine(cfg.Workspace, "PLUMA.md"));
                AppendFile(sb, Path.Combine(cfg.Workspace, "context", "estilo.md"));
            }
            sb.Append(user);
            return sb.ToString();
        }

        static void AppendFile(StringBuilder sb, string path)
        {
            if (!File.Exists(path)) return;
            string text = File.ReadAllText(path, Encoding.UTF8).Trim();
            if (text.Length == 0) return;
            sb.Append(text).Append("\n\n");
        }

        static string BuildPrompt(string instr, string texto, int imageCount)
        {
            texto = (texto ?? "").Trim();
            string adjuntas = imageCount == 0 ? ""
                : imageCount == 1 ? "(Te envío una imagen.)" : "(Te envío " + imageCount + " imágenes.)";

            if (string.IsNullOrEmpty(instr))
            {
                if (texto.Length == 0) return adjuntas + " Mirala y respondé a lo que plantea.";
                return adjuntas.Length > 0 ? adjuntas + "\n\n" + texto : texto;
            }
            string msg = "Instrucción: " + instr;
            if (adjuntas.Length > 0) msg += "\n" + adjuntas;
            if (texto.Length > 0) msg += "\n\n<texto>\n" + texto + "\n</texto>";
            return msg;
        }

        static string BuildClaudeMessage(string text, List<Attachment> images)
        {
            var content = new List<object>();
            foreach (Attachment img in images)
            {
                content.Add(new Dictionary<string, object> {
                    { "type", "image" },
                    { "source", new Dictionary<string, object> {
                        { "type", "base64" }, { "media_type", img.MediaType }, { "data", img.Base64 } } }
                });
            }
            content.Add(new Dictionary<string, object> {
                { "type", "text" }, { "text", text } });

            var line = new Dictionary<string, object> {
                { "type", "user" },
                { "message", new Dictionary<string, object> { { "role", "user" }, { "content", content } } }
            };
            var json = new JavaScriptSerializer();
            json.MaxJsonLength = int.MaxValue;
            return json.Serialize(line) + "\n";
        }

        static Launch BuildLaunch(Agent agent, Config cfg, string baseDir, string command,
            string textPrompt, List<string> imagePaths, string runDir)
        {
            string rulesPath = Path.Combine(baseDir, "word-rules.md");
            var args = new List<string>();
            foreach (string a in agent.Args) args.Add(Subst(a, cfg, runDir, rulesPath));

            if (agent.Rules == "file-arg" && !string.IsNullOrEmpty(agent.RulesArg))
            {
                args.Add(agent.RulesArg);
                args.Add(rulesPath);
            }

            string model = ResolveModel(agent, cfg);
            if (model.Length > 0 && !string.IsNullOrEmpty(agent.ModelArg))
            {
                args.Add(agent.ModelArg);
                args.Add(model);
            }

            if (agent.Images == "files" && imagePaths != null)
            {
                foreach (string path in imagePaths)
                {
                    args.Add(agent.ImageArg);
                    args.Add(path);
                }
            }

            if (agent.PromptVia == "arg")
            {
                if (!string.IsNullOrEmpty(agent.PromptArg)) args.Add(agent.PromptArg);
                args.Add(textPrompt ?? "");
            }
            else if (agent.PromptVia == "file")
            {
                args.Add(agent.PromptArg);
                args.Add(Path.Combine(runDir, "prompt.txt"));
            }

            var launch = new Launch();
            launch.FileName = command;
            launch.ArgList = args;
            launch.Arguments = JoinArgs(args);
            launch.Env = ResolveEnv(agent, cfg, runDir, rulesPath);
            launch.WorkDir = cfg.Workspace;
            ApplyCmdShim(launch, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "cmd.exe"));
            return launch;
        }

        static void ApplyCmdShim(Launch launch, string cmdExe)
        {
            if (!IsCmdScript(launch.FileName)) return;
            string inner = QuoteArg(launch.FileName, true);
            if (launch.Arguments.Length > 0) inner += " " + launch.Arguments;
            launch.Arguments = "/d /s /c \"" + inner + "\"";
            launch.FileName = cmdExe;
        }

        static string JoinArgs(List<string> args)
        {
            var sb = new StringBuilder();
            foreach (string a in args)
            {
                if (sb.Length > 0) sb.Append(' ');
                sb.Append(QuoteArg(a, false));
            }
            return sb.ToString();
        }

        // Quoting de CreateProcess: https://learn.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-commandlinetoargvw
        static string QuoteArg(string s, bool force)
        {
            if (s == null) s = "";
            if (s.Length == 0) return "\"\"";
            bool need = force;
            if (!need)
            {
                foreach (char c in s)
                {
                    if (c == ' ' || c == '\t' || c == '"' || c == '\n' || c == '\r') { need = true; break; }
                }
            }
            if (!need) return s;
            var sb = new StringBuilder();
            sb.Append('"');
            int bs = 0;
            foreach (char c in s)
            {
                if (c == '\\') { bs++; continue; }
                if (c == '"')
                {
                    sb.Append('\\', bs * 2 + 1);
                    sb.Append('"');
                    bs = 0;
                    continue;
                }
                if (bs > 0) { sb.Append('\\', bs); bs = 0; }
                if (c == '\n' || c == '\r' || c == '\t') sb.Append(' ');
                else sb.Append(c);
            }
            if (bs > 0) sb.Append('\\', bs * 2);
            sb.Append('"');
            return sb.ToString();
        }

        static Process StartProcess(Launch launch, string stdinPayload, Agent agent)
        {
            Log(launch.FileName + " " + Redact(launch.Arguments, stdinPayload));
            Log("prompt via " + agent.PromptVia + " (" + (stdinPayload != null ? stdinPayload.Length : 0) + " chars en stdin)");

            var psi = new ProcessStartInfo(launch.FileName, launch.Arguments);
            psi.WorkingDirectory = launch.WorkDir;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.RedirectStandardInput = true;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;
            psi.StandardOutputEncoding = Encoding.UTF8;
            psi.StandardErrorEncoding = Encoding.UTF8;
            foreach (KeyValuePair<string, string> kv in launch.Env)
                psi.EnvironmentVariables[kv.Key] = kv.Value;

            Process p = Process.Start(psi);
            p.ErrorDataReceived += delegate(object s, DataReceivedEventArgs e)
            {
                if (e.Data != null) lock (Gate) StdErr.AppendLine(e.Data);
            };
            p.BeginErrorReadLine();

            var reader = new Thread(delegate()
            {
                try
                {
                    if (agent.Parser == "text") ReadText(p, agent.Stream);
                    else ReadLines(p, agent);
                }
                catch (Exception ex)
                {
                    lock (Gate) { if (error == null) error = ex.Message; }
                }
                lock (Gate) { finished = true; }
            });
            reader.IsBackground = true;
            reader.Start();

            if (stdinPayload != null)
            {
                byte[] bytes = new UTF8Encoding(false).GetBytes(stdinPayload);
                p.StandardInput.BaseStream.Write(bytes, 0, bytes.Length);
            }
            p.StandardInput.Close();
            return p;
        }

        static string Redact(string arguments, string secret)
        {
            if (string.IsNullOrEmpty(secret) || secret.Length < 24) return arguments;
            string quoted = QuoteArg(secret, false);
            if (arguments.IndexOf(quoted, StringComparison.Ordinal) >= 0)
                return arguments.Replace(quoted, "\"…\"");
            return arguments;
        }

        static void ReadLines(Process p, Agent agent)
        {
            var json = new JavaScriptSerializer();
            json.MaxJsonLength = int.MaxValue;
            string line;
            while ((line = p.StandardOutput.ReadLine()) != null)
            {
                if (agent.Parser == "claude-stream") HandleClaude(json, line);
                else ApplyJsonl(agent, json, line);
            }
        }

        static void ReadText(Process p, bool stream)
        {
            var buf = new char[512];
            var held = new StringBuilder();
            int n;
            while ((n = p.StandardOutput.Read(buf, 0, buf.Length)) > 0)
            {
                string chunk = StripNoise(new string(buf, 0, n));
                if (chunk.Length == 0) continue;
                if (stream) lock (Gate) Answer.Append(chunk);
                else held.Append(chunk);
            }
            if (!stream)
            {
                string all = held.ToString().Trim();
                lock (Gate) { if (all.Length > 0) resultText = all; }
            }
        }

        static string StripNoise(string s)
        {
            s = Regex.Replace(s, "\x1b\\[[0-9;?]*[ -/]*[@-~]", "");
            return s.Replace("\r", "");
        }

        // Eventos de "claude -p --output-format stream-json --include-partial-messages".
        // El texto anterior a una herramienta se descarta: queda lo posterior a la ultima.
        static void HandleClaude(JavaScriptSerializer json, string line)
        {
            Dictionary<string, object> o;
            try { o = json.DeserializeObject(line) as Dictionary<string, object>; }
            catch { return; }
            if (o == null) return;

            string type = Str(o, "type");
            if (type == "stream_event")
            {
                var ev = Dict(o, "event");
                if (ev == null) return;
                string et = Str(ev, "type");
                if (et == "content_block_start")
                {
                    var block = Dict(ev, "content_block");
                    if (block != null && Str(block, "type") == "tool_use")
                        lock (Gate) { Answer.Clear(); status = StatusReadingClaude; }
                }
                else if (et == "content_block_delta")
                {
                    var delta = Dict(ev, "delta");
                    if (delta != null && Str(delta, "type") == "text_delta")
                        lock (Gate) Answer.Append(Str(delta, "text"));
                }
            }
            else if (type == "result")
            {
                object isErr;
                bool failed = (o.TryGetValue("is_error", out isErr) && isErr is bool && (bool)isErr)
                              || Str(o, "subtype") != "success";
                lock (Gate)
                {
                    resultText = Str(o, "result");
                    if (failed) error = resultText ?? Str(o, "subtype");
                }
            }
        }

        static void ApplyJsonl(Agent agent, JavaScriptSerializer json, string line)
        {
            Dictionary<string, object> o;
            try { o = json.DeserializeObject(line) as Dictionary<string, object>; }
            catch { return; }
            if (o == null) return;
            string type = Str(o, "type");

            foreach (JsonlRule rule in agent.TextRules)
            {
                if (!RuleMatches(rule, o, type)) continue;
                string val = StringAt(o, rule.Path);
                if (string.IsNullOrEmpty(val)) continue;
                lock (Gate) ApplyText(rule, o, val);
            }
            foreach (JsonlRule rule in agent.ErrorRules)
            {
                if (!RuleMatches(rule, o, type)) continue;
                // Con path, solo es error si el campo vino con texto. Si no, el
                // evento puede ser un result exitoso que no trae "error".
                if (string.IsNullOrEmpty(rule.Path))
                {
                    lock (Gate) { error = "el agente informó un error"; }
                }
                else
                {
                    string val = StringAt(o, rule.Path);
                    if (!string.IsNullOrEmpty(val)) lock (Gate) { error = val; }
                }
            }
            foreach (JsonlRule rule in agent.StatusRules)
            {
                if (!RuleMatches(rule, o, type)) continue;
                lock (Gate)
                {
                    status = Reading(agent);
                    if (rule.Clear) Answer.Clear();
                }
            }
            ApplyParserTail(agent, o, type);
        }

        // Lo que el matcher por igualdad no puede expresar: códigos de salida
        // y stop reasons que no son "end_turn".
        static void ApplyParserTail(Agent agent, Dictionary<string, object> o, string type)
        {
            if (agent.Parser == "grok-stream" && type == "end")
            {
                string stop = Str(o, "stopReason");
                if (stop == "refusal" || stop == "cancelled" || stop == "max_turn_requests")
                    lock (Gate) { if (error == null) error = stop; }
            }
            else if (agent.Parser == "hermes-stream" && type == "result")
            {
                string errMsg = StringAt(o, "error");
                object code = At(o, "exit_code");
                if (code != null)
                {
                    int exit = 0;
                    try { exit = Convert.ToInt32(code, CultureInfo.InvariantCulture); }
                    catch { exit = 0; }
                    if (exit != 0 && string.IsNullOrEmpty(errMsg))
                        errMsg = "terminó con código " + exit.ToString(CultureInfo.InvariantCulture);
                }
                if (!string.IsNullOrEmpty(errMsg)) lock (Gate) { error = errMsg; }
            }
        }

        static void ApplyText(JsonlRule rule, Dictionary<string, object> o, string val)
        {
            string mode = rule.Mode ?? "append";
            if (mode == "set")
            {
                Answer.Clear();
                Answer.Append(val);
            }
            else if (mode == "set-if-empty")
            {
                if (Answer.Length == 0) Answer.Append(val);
            }
            else if (mode == "auto")
            {
                object delta = At(o, "delta");
                if (delta is bool && (bool)delta) Answer.Append(val);
                else { Answer.Clear(); Answer.Append(val); }
            }
            else Answer.Append(val);
        }

        static bool RuleMatches(JsonlRule rule, Dictionary<string, object> o, string type)
        {
            if (!string.IsNullOrEmpty(rule.Type) && rule.Type != type) return false;
            foreach (KeyValuePair<string, object> kv in rule.Match)
            {
                if (!Same(kv.Value, At(o, kv.Key))) return false;
            }
            return true;
        }

        static bool Same(object expected, object actual)
        {
            if (expected == null || actual == null) return expected == actual;
            bool eb, ab;
            if (AsBool(expected, out eb) && AsBool(actual, out ab)) return eb == ab;
            return string.Equals(Convert.ToString(expected, CultureInfo.InvariantCulture),
                                 Convert.ToString(actual, CultureInfo.InvariantCulture),
                                 StringComparison.Ordinal);
        }

        static bool AsBool(object v, out bool b)
        {
            if (v is bool) { b = (bool)v; return true; }
            string s = v as string;
            if (s == "true") { b = true; return true; }
            if (s == "false") { b = false; return true; }
            b = false;
            return false;
        }

        static object At(Dictionary<string, object> o, string path)
        {
            if (o == null || string.IsNullOrEmpty(path)) return null;
            object cur = o;
            foreach (string part in path.Split('.'))
            {
                var d = cur as Dictionary<string, object>;
                if (d == null || !d.TryGetValue(part, out cur)) return null;
            }
            return cur;
        }

        static string StringAt(Dictionary<string, object> o, string path)
        {
            if (string.IsNullOrEmpty(path)) return null;
            return At(o, path) as string;
        }

        static string Thinking(Agent agent)
        {
            return "✳ " + agent.Label + " está pensando…";
        }

        static string Reading(Agent agent)
        {
            if (agent != null && agent.Id == "claude") return StatusReadingClaude;
            string label = agent == null ? "Pluma" : agent.Label;
            return "✳ " + label + " está leyendo el contexto…";
        }

        static string LookupTask(string workspace, string taskId)
        {
            taskId = (taskId ?? "").Trim().ToLowerInvariant();
            if (taskId.Length == 0) return null;
            string path = string.IsNullOrEmpty(workspace) ? null : Path.Combine(workspace, "tasks.json");
            if (path != null && File.Exists(path))
            {
                try
                {
                    var json = new JavaScriptSerializer();
                    var o = json.DeserializeObject(File.ReadAllText(path, Encoding.UTF8)) as Dictionary<string, object>;
                    object v;
                    if (o != null && o.TryGetValue(taskId, out v))
                    {
                        string s = v as string;
                        if (!string.IsNullOrEmpty(s) && s.Trim().Length > 0) return s.Trim();
                    }
                }
                catch { }
            }
            string built;
            if (DefaultTasks().TryGetValue(taskId, out built)) return built;
            return null;
        }

        static Dictionary<string, string> DefaultTasks()
        {
            var d = new Dictionary<string, string>();
            d["corregir"] = "Corregí la ortografía, la gramática y la puntuación. No cambies el tono ni el significado.";
            d["formal"] = "Reescribilo en un registro más formal, sin cambiar el significado.";
            d["resumir"] = "Resumilo en menos palabras, conservando las ideas principales.";
            d["expandir"] = "Expandilo con más detalle y claridad, sin inventar datos que no estén en el texto.";
            d["traducir"] = "Traducilo: si está en español, al inglés; si está en otro idioma, al español rioplatense. Conservá el significado.";
            return d;
        }

        static string Clean(string s)
        {
            if (s == null) return "";
            s = s.Replace("\r\n", "\n");
            s = Regex.Replace(s, @"(?m)^#{1,6}\s+", "");
            s = Regex.Replace(s, @"\*\*(.+?)\*\*", "$1");
            s = Regex.Replace(s, @"\n{2,}", "\n");
            return s.Trim().Replace("\n", "\r");
        }

        static void ResetRun()
        {
            lock (Gate)
            {
                Answer.Clear();
                StdErr.Clear();
                status = null;
                resultText = null;
                error = null;
                finished = false;
            }
        }

        static void Kill(Process p)
        {
            try
            {
                if (p == null || p.HasExited) return;
                var psi = new ProcessStartInfo();
                psi.FileName = "taskkill";
                psi.Arguments = "/F /T /PID " + p.Id.ToString(CultureInfo.InvariantCulture);
                psi.CreateNoWindow = true;
                psi.UseShellExecute = false;
                Process killer = Process.Start(psi);
                if (killer != null) killer.WaitForExit(3000);
            }
            catch
            {
                try { if (p != null && !p.HasExited) p.Kill(); } catch { }
            }
        }

        static string BookmarkName(string tag) { return "_" + tag.Replace("-", "_"); }

        // ---------- Word (COM) ----------

        static dynamic FindDocument(dynamic word, string tag)
        {
            int count = word.Documents.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic d = word.Documents.Item(i);
                if (d.SelectContentControlsByTag(tag).Count > 0) return d;
            }
            return null;
        }

        static dynamic GetBlock(dynamic doc, string tag)
        {
            dynamic found = doc.SelectContentControlsByTag(tag);
            return found.Count > 0 ? found.Item(1) : null;
        }

        static void WriteBlock(dynamic cc, string text, int color)
        {
            cc.LockContents = false;
            cc.Range.Text = text;
            cc.Range.Font.Color = color;
            cc.LockContents = true;
        }

        static void AppendBlock(dynamic cc, string suffix)
        {
            cc.LockContents = false;
            cc.Range.InsertAfter(suffix);
            cc.LockContents = true;
        }

        static void Finish(dynamic doc, string tag, string text, int color)
        {
            try
            {
                dynamic cc = GetBlock(doc, tag);
                if (cc != null)
                {
                    if (text != null) WriteBlock(cc, text, color);
                    cc.LockContents = false;
                    try { doc.Bookmarks.Add("_pluma_r" + tag.Substring(TagPrefix.Length), cc.Range); }
                    catch (Exception ex) { Log("marcador de respuesta: " + ex.Message); }
                    cc.Delete(false);
                }
            }
            catch (Exception ex) { Log("finish: " + ex.Message); }
            DeleteVariables(doc, tag);
            try { doc.Application.StatusBar = " "; } catch { }
        }

        static void FinishEmpty(dynamic doc, string tag)
        {
            try
            {
                dynamic cc = GetBlock(doc, tag);
                if (cc != null)
                {
                    dynamic par = cc.Range.Paragraphs.First.Range;
                    cc.LockContents = false;
                    cc.Delete(true);
                    string rest = par.Text;
                    if (rest.Trim('\r', ' ').Length == 0) par.Delete();
                }
            }
            catch (Exception ex) { Log("finish vacio: " + ex.Message); }
            DeleteVariables(doc, tag);
            try { doc.Application.StatusBar = " "; } catch { }
        }

        static void SetVariable(dynamic doc, string name, string value)
        {
            try { doc.Variables.Item(name).Value = value; }
            catch
            {
                try { doc.Variables.Add(name, value); } catch { }
            }
        }

        static string ReadVariable(dynamic doc, string name)
        {
            try { return (string)doc.Variables.Item(name).Value; }
            catch { return null; }
        }

        static void DeleteVariables(dynamic doc, string tag)
        {
            foreach (string suffix in new[] { "-text", "-instr", "-color", "-stop", "-hb", "-agent", "-task" })
            {
                try { doc.Variables.Item(tag + suffix).Delete(); } catch { }
            }
            try
            {
                if (doc.Bookmarks.Exists(BookmarkName(tag))) doc.Bookmarks.Item(BookmarkName(tag)).Delete();
            }
            catch { }
        }

        static string Str(Dictionary<string, object> o, string key)
        {
            object v;
            return o.TryGetValue(key, out v) ? v as string : null;
        }

        static Dictionary<string, object> Dict(Dictionary<string, object> o, string key)
        {
            object v;
            return o.TryGetValue(key, out v) ? v as Dictionary<string, object> : null;
        }

        internal static void Log(string msg)
        {
            try { File.AppendAllText(logPath, DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg + "\r\n"); }
            catch { }
        }

        // ---------- self-test (sin Word) ----------

        static int SelfTest()
        {
            int failed = 0;
            failed += Expect("prompt pelado", BuildPrompt(null, "hola", 0), "hola");
            failed += Expect("prompt imagen", BuildPrompt(null, "", 1), "(Te envío una imagen.) Mirala y respondé a lo que plantea.");
            failed += Expect("prompt instruccion", BuildPrompt("formal", "hola", 0), "Instrucción: formal\n\n<texto>\nhola\n</texto>");
            failed += Expect("prompt dos imagenes", BuildPrompt("formal", "", 2), "Instrucción: formal\n(Te envío 2 imágenes.)");
            failed += Expect("clean markdown", Clean("## Titulo\n\n**hola**\n\n\nchau"), "Titulo\rhola\rchau");

            string quoted = QuoteArg("a b", false);
            failed += Expect("quote espacio", quoted, "\"a b\"");
            failed += Expect("quote comilla", QuoteArg("say \"hi\"", false), "\"say \\\"hi\\\"\"");
            failed += Expect("quote vacio", QuoteArg("", false), "\"\"");
            failed += Expect("quote barra final", QuoteArg("a b\\", false), "\"a b\\\\\"");
            failed += Expect("quote simple", QuoteArg("Read,Glob,Grep,Skill", false), "Read,Glob,Grep,Skill");

            var args = new List<string>();
            args.Add("--tools");
            args.Add("Read,Glob,Grep,Skill");
            args.Add("--model");
            args.Add("sonnet");
            failed += Expect("join", JoinArgs(args), "--tools Read,Glob,Grep,Skill --model sonnet");

            var shim = new Launch();
            shim.FileName = "C:\\npm\\codex.cmd";
            shim.Arguments = "exec - --json";
            ApplyCmdShim(shim, "C:\\Windows\\System32\\cmd.exe");
            failed += Expect("shim exe", shim.FileName, "C:\\Windows\\System32\\cmd.exe");
            failed += Expect("shim args", shim.Arguments, "/d /s /c \"\"C:\\npm\\codex.cmd\" exec - --json\"");

            failed += Expect("arg largo", PromptProblem(new Agent { PromptVia = "arg" }, new string('x', 7001)) == null ? "no" : "si", "si");
            failed += Expect("file sin flag", PromptProblem(new Agent { PromptVia = "file" }, "hola") == null ? "no" : "si", "si");

            ResetRun();
            var json = new JavaScriptSerializer();
            json.MaxJsonLength = int.MaxValue;
            HandleClaude(json, "{\"type\":\"stream_event\",\"event\":{\"type\":\"content_block_delta\",\"delta\":{\"type\":\"text_delta\",\"text\":\"Hola\"}}}");
            HandleClaude(json, "{\"type\":\"stream_event\",\"event\":{\"type\":\"content_block_start\",\"content_block\":{\"type\":\"tool_use\",\"name\":\"Read\"}}}");
            failed += Expect("claude limpia al usar herramienta", Answer.ToString(), "");
            failed += Expect("claude estado", status, StatusReadingClaude);
            HandleClaude(json, "{\"type\":\"stream_event\",\"event\":{\"type\":\"content_block_delta\",\"delta\":{\"type\":\"text_delta\",\"text\":\"Chau\"}}}");
            HandleClaude(json, "{\"type\":\"result\",\"subtype\":\"success\",\"result\":\"Chau\",\"is_error\":false}");
            failed += Expect("claude texto", Answer.ToString(), "Chau");
            failed += Expect("claude result", resultText, "Chau");
            failed += Expect("claude sin error", error == null ? "-" : error, "-");

            ResetRun();
            HandleClaude(json, "{\"type\":\"result\",\"subtype\":\"error_during_execution\",\"is_error\":true,\"result\":\"login\"}");
            failed += Expect("claude error", error, "login");

            string claudeLine = BuildClaudeMessage("hola", new List<Attachment>());
            var parsed = json.DeserializeObject(claudeLine) as Dictionary<string, object>;
            var message = Dict(parsed, "message");
            failed += Expect("claude mensaje rol", Str(message, "role"), "user");
            object first = null;
            foreach (object item in (IEnumerable)message["content"]) { first = item; break; }
            var block = first as Dictionary<string, object>;
            failed += Expect("claude mensaje texto", Str(block, "text"), "hola");

            Agent codex = new Agent();
            codex.Id = "codex";
            codex.Label = "Codex";
            codex.Parser = "codex-jsonl";
            AgentFile.ApplyDefaults(codex);
            ResetRun();
            ApplyJsonl(codex, json, "{\"type\":\"item.started\",\"item\":{\"type\":\"command_execution\",\"command\":\"rg\"}}");
            failed += Expect("codex estado", status, "✳ Codex está leyendo el contexto…");
            ApplyJsonl(codex, json, "{\"type\":\"item.completed\",\"item\":{\"type\":\"agent_message\",\"text\":\"Listo.\"}}");
            failed += Expect("codex texto", Answer.ToString(), "Listo.");
            ResetRun();
            ApplyJsonl(codex, json, "{\"type\":\"turn.failed\",\"error\":{\"message\":\"boom\"}}");
            failed += Expect("codex error", error, "boom");

            Agent gemini = new Agent();
            gemini.Id = "gemini";
            gemini.Label = "Gemini";
            gemini.Parser = "gemini-stream";
            AgentFile.ApplyDefaults(gemini);
            ResetRun();
            ApplyJsonl(gemini, json, "{\"type\":\"message\",\"role\":\"assistant\",\"content\":\"Hola\",\"delta\":true}");
            ApplyJsonl(gemini, json, "{\"type\":\"message\",\"role\":\"assistant\",\"content\":\" mundo\",\"delta\":true}");
            ApplyJsonl(gemini, json, "{\"type\":\"tool_use\",\"tool_name\":\"read\"}");
            failed += Expect("gemini estado", status, "✳ Gemini está leyendo el contexto…");
            ApplyJsonl(gemini, json, "{\"type\":\"message\",\"role\":\"assistant\",\"content\":\"Listo\",\"delta\":true}");
            failed += Expect("gemini deltas", Answer.ToString(), "Hola mundoListo");
            ResetRun();
            ApplyJsonl(gemini, json, "{\"type\":\"result\",\"status\":\"success\",\"response\":\"Solo esto\"}");
            failed += Expect("gemini result", Answer.ToString(), "Solo esto");
            ResetRun();
            ApplyJsonl(gemini, json, "{\"type\":\"result\",\"status\":\"error\",\"error\":{\"message\":\"quota\"}}");
            failed += Expect("gemini error", error, "quota");

            ResetRun();
            string noisy = StripNoise("\u001b[31mHola\u001b[0m\r\nmundo");
            failed += Expect("ansi", noisy, "Hola\nmundo");

            string dir = Path.Combine(Path.GetTempPath(), "pluma-test-" + Guid.NewGuid().ToString("n"));
            Directory.CreateDirectory(dir);
            try
            {
                File.WriteAllText(Path.Combine(dir, "tasks.json"), "{\"corregir\":\"Custom\"}", new UTF8Encoding(false));
                failed += Expect("tarea archivo", LookupTask(dir, "corregir"), "Custom");
                failed += Expect("tarea builtin", LookupTask(dir, "formal").IndexOf("formal", StringComparison.Ordinal) >= 0 ? "si" : "no", "si");
                failed += Expect("tarea desconocida", LookupTask(dir, "no-existe") == null ? "-" : "x", "-");

                var cfg = new Config();
                cfg.Workspace = dir;
                cfg.Model = "sonnet";
                cfg.ClaudeConfigDir = Path.Combine(dir, "claude-home");
                cfg.ClaudePath = Path.Combine(dir, "claude.exe");
                File.WriteAllText(cfg.ClaudePath, "");
                File.WriteAllText(Path.Combine(dir, "word-rules.md"), "REGLAS", new UTF8Encoding(false));
                File.WriteAllText(Path.Combine(dir, "PLUMA.md"), "ROL", new UTF8Encoding(false));

                Agent claude = AgentFile.Builtin(cfg).Resolve(null);
                failed += Expect("builtin id", claude.Id, "claude");
                failed += Expect("builtin via", claude.PromptVia, "stdin");
                failed += Expect("builtin format", claude.PromptFormat, "claude-message");
                failed += Expect("builtin parser", claude.Parser, "claude-stream");
                failed += Expect("builtin images", claude.Images, "claude");
                failed += Expect("builtin args", string.Join(" ", claude.Args.ToArray()), string.Join(" ", ClaudeArgs()));
                string composed = ComposePrompt(claude, cfg, dir, null, "hola", 0);
                failed += Expect("claude sin preambulo", composed, "hola");

                Launch launch = BuildLaunch(claude, cfg, dir, cfg.ClaudePath, composed, new List<string>(), null);
                failed += Expect("launch model", launch.Arguments.IndexOf("--model sonnet", StringComparison.Ordinal) >= 0 ? "si" : "no", "si");
                failed += Expect("launch rules", launch.Arguments.IndexOf("--append-system-prompt-file", StringComparison.Ordinal) >= 0 ? "si" : "no", "si");
                failed += Expect("launch tools", launch.Arguments.IndexOf("--tools Read,Glob,Grep,Skill", StringComparison.Ordinal) >= 0 ? "si" : "no", "si");
                failed += Expect("launch env", launch.Env.ContainsKey("CLAUDE_CONFIG_DIR") ? launch.Env["CLAUDE_CONFIG_DIR"] : "-", cfg.ClaudeConfigDir);

                Agent otro = new Agent();
                otro.Id = "otro";
                otro.Label = "Otro";
                otro.Command = "otro";
                otro.Rules = "prepend";
                otro.Context = "prepend";
                otro.PromptVia = "stdin";
                string full = ComposePrompt(otro, cfg, dir, "resumí", "texto", 0);
                failed += Expect("prepend", full.IndexOf("REGLAS", StringComparison.Ordinal) >= 0 && full.IndexOf("ROL", StringComparison.Ordinal) >= 0 && full.IndexOf("<texto>", StringComparison.Ordinal) >= 0 ? "si" : "no", "si");

                string template = FindUp(Directory.GetCurrentDirectory(), Path.Combine("workspace-template", "agents.json"));
                if (template == null) template = FindUp(AppDomain.CurrentDomain.BaseDirectory, Path.Combine("workspace-template", "agents.json"));
                if (template == null)
                {
                    Console.WriteLine("ok  template ausente, salteo");
                }
                else
                {
                    var bare = new Config();
                    bare.Workspace = Path.GetDirectoryName(Path.GetDirectoryName(template));
                    AgentFile file = AgentFile.Load(template, bare);
                    failed += Expect("template default", file.DefaultId, "claude");
                    Agent tClaude = file.Resolve("claude");
                    failed += Expect("template claude args", string.Join(" ", tClaude.Args.ToArray()), string.Join(" ", ClaudeArgs()));
                    failed += Expect("template claude parser", tClaude.Parser, "claude-stream");
                    failed += Expect("template claude images", tClaude.Images, "claude");
                    failed += Expect("template claude context", tClaude.Context, "none");
                    Agent tCodex = file.Resolve("codex");
                    failed += Expect("template codex via", tCodex.PromptVia + " " + tCodex.Args[0] + " " + tCodex.Args[1], "stdin exec -");
                    failed += Expect("template codex images", tCodex.Images, "files");
                    ResetRun();
                    ApplyJsonl(tCodex, json, "{\"type\":\"item.completed\",\"item\":{\"type\":\"agent_message\",\"text\":\"Desde plantilla\"}}");
                    failed += Expect("template codex parse", Answer.ToString(), "Desde plantilla");
                    Agent tGemini = file.Resolve("gemini");
                    ResetRun();
                    ApplyJsonl(tGemini, json, "{\"type\":\"message\",\"role\":\"assistant\",\"content\":\"Hi\",\"delta\":true}");
                    failed += Expect("template gemini parse", Answer.ToString(), "Hi");
                    Agent tAider = file.Resolve("aider");
                    failed += Expect("template aider", tAider.PromptVia + " " + tAider.PromptArg + " " + (tAider.Stream ? "stream" : "bloque"), "file --message-file bloque");
                    Agent tGrok = file.Resolve("grok");
                    failed += Expect("template grok", tGrok.PromptVia + " " + tGrok.PromptArg + " " + tGrok.Parser + " " + tGrok.Images, "file --prompt-file grok-stream none");
                    ResetRun();
                    ApplyJsonl(tGrok, json, "{\"type\":\"text\",\"data\":\"Hola\"}");
                    ApplyJsonl(tGrok, json, "{\"type\":\"tool_call\",\"toolName\":\"read_file\",\"status\":\"in_progress\"}");
                    failed += Expect("grok limpia", Answer.ToString(), "");
                    failed += Expect("grok estado", status, "✳ Grok está leyendo el contexto…");
                    ApplyJsonl(tGrok, json, "{\"type\":\"text\",\"data\":\"Ok\"}");
                    ApplyJsonl(tGrok, json, "{\"type\":\"end\",\"stopReason\":\"end_turn\"}");
                    failed += Expect("grok texto", Answer.ToString(), "Ok");
                    failed += Expect("grok sin error", error == null ? "-" : error, "-");
                    ResetRun();
                    ApplyJsonl(tGrok, json, "{\"type\":\"error\",\"message\":\"no auth\"}");
                    failed += Expect("grok error", error, "no auth");
                    Agent tHermes = file.Resolve("hermes");
                    failed += Expect("template hermes", tHermes.Parser + " " + tHermes.Images + " " + tHermes.MaxImages.ToString(CultureInfo.InvariantCulture), "hermes-stream files 1");
                    ResetRun();
                    ApplyJsonl(tHermes, json, "{\"type\":\"text\",\"text\":\"Ho\"}");
                    ApplyJsonl(tHermes, json, "{\"type\":\"tool_use\",\"name\":\"read_file\"}");
                    failed += Expect("hermes limpia", Answer.ToString(), "");
                    ApplyJsonl(tHermes, json, "{\"type\":\"text\",\"text\":\"Fi\"}");
                    ApplyJsonl(tHermes, json, "{\"type\":\"result\",\"exit_code\":0,\"text\":\"Fin\"}");
                    failed += Expect("hermes texto final", Answer.ToString(), "Fin");
                    failed += Expect("hermes sin error", error == null ? "-" : error, "-");
                    ResetRun();
                    ApplyJsonl(tHermes, json, "{\"type\":\"result\",\"exit_code\":1,\"error\":\"falló\"}");
                    failed += Expect("hermes error", error, "falló");
                    string tasks = FindUp(Directory.GetCurrentDirectory(), Path.Combine("workspace-template", "tasks.json"));
                    if (tasks != null)
                    {
                        foreach (KeyValuePair<string, string> kv in DefaultTasks())
                        {
                            string got = LookupTask(Path.GetDirectoryName(tasks), kv.Key);
                            failed += Expect("tasks " + kv.Key, got, kv.Value);
                        }
                    }
                }
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch { }
            }

            if (failed == 0)
            {
                Console.WriteLine("self-test ok");
                return 0;
            }
            Console.WriteLine("self-test FAIL " + failed);
            return 1;
        }

        static int Expect(string name, string got, string want)
        {
            if (got == want)
            {
                Console.WriteLine("ok  " + name);
                return 0;
            }
            Console.WriteLine("FAIL " + name);
            Console.WriteLine("  want: " + Show(want));
            Console.WriteLine("  got:  " + Show(got));
            return 1;
        }

        static string Show(string s)
        {
            if (s == null) return "<null>";
            return s.Replace("\r", "\\r").Replace("\n", "\\n");
        }

        static string FindUp(string start, string relative)
        {
            try
            {
                var dir = new DirectoryInfo(start);
                while (dir != null)
                {
                    string path = Path.Combine(dir.FullName, relative);
                    if (File.Exists(path)) return path;
                    dir = dir.Parent;
                }
            }
            catch { }
            return null;
        }

        // ---------- COM: reintentar cuando Word esta ocupado ----------

        [DllImport("ole32.dll")]
        static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        class RetryFilter : IOleMessageFilter
        {
            public int HandleInComingCall(int callType, IntPtr caller, int tickCount, IntPtr info) { return 0; }

            public int RetryRejectedCall(IntPtr callee, int tickCount, int rejectType)
            {
                const int RetryLater = 2;
                if (rejectType == RetryLater && tickCount < 60000) return 100;
                return -1;
            }

            public int MessagePending(IntPtr callee, int tickCount, int pendingType) { return 2; }
        }
    }

    [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int callType, IntPtr caller, int tickCount, IntPtr info);
        [PreserveSig] int RetryRejectedCall(IntPtr callee, int tickCount, int rejectType);
        [PreserveSig] int MessagePending(IntPtr callee, int tickCount, int pendingType);
    }

    class Launch
    {
        public string FileName;
        public string Arguments;
        public List<string> ArgList;
        public Dictionary<string, string> Env;
        public string WorkDir;
    }

    class JsonlRule
    {
        public string Type;
        public string Path;
        public string Mode;
        public bool Clear;
        public Dictionary<string, object> Match = new Dictionary<string, object>();
    }

    class Agent
    {
        public string Id;
        public string Label;
        public bool Enabled = true;
        public string Command;
        public List<string> Args = new List<string>();
        public string PromptVia = "stdin";
        public string PromptFormat = "text";
        public string PromptArg;
        public string Parser = "text";
        public bool Stream = true;
        public string Images = "none";
        public string ImageArg = "-i";
        public int MaxImages;
        public string Model;
        public string ModelArg;
        public string Rules = "none";
        public string RulesArg = "--append-system-prompt-file";
        public string Context = "none";
        public Dictionary<string, string> Env = new Dictionary<string, string>();
        public List<JsonlRule> TextRules = new List<JsonlRule>();
        public List<JsonlRule> ErrorRules = new List<JsonlRule>();
        public List<JsonlRule> StatusRules = new List<JsonlRule>();
    }

    class AgentFile
    {
        public string DefaultId = "claude";
        public List<Agent> Agents = new List<Agent>();

        public static AgentFile Builtin(Config cfg)
        {
            var file = new AgentFile();
            var a = new Agent();
            a.Id = "claude";
            a.Label = "Claude";
            a.Enabled = true;
            a.Command = cfg.ClaudePath ?? "";
            a.Args = new List<string>(ProgramClaudeArgs());
            a.PromptVia = "stdin";
            a.PromptFormat = "claude-message";
            a.Parser = "claude-stream";
            a.Images = "claude";
            a.ModelArg = "--model";
            a.Rules = "file-arg";
            a.RulesArg = "--append-system-prompt-file";
            a.Context = "none";
            a.Env["CLAUDE_CONFIG_DIR"] = "{claudeConfigDir}";
            file.Agents.Add(a);
            return file;
        }

        static string[] ProgramClaudeArgs()
        {
            return new[] {
                "-p",
                "--input-format", "stream-json",
                "--output-format", "stream-json",
                "--include-partial-messages",
                "--verbose",
                "--setting-sources", "project",
                "--strict-mcp-config",
                "--no-session-persistence",
                "--tools", "Read,Glob,Grep,Skill"
            };
        }

        public static AgentFile Load(string path, Config cfg)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return Builtin(cfg);
            var json = new JavaScriptSerializer();
            json.MaxJsonLength = int.MaxValue;
            var root = json.DeserializeObject(File.ReadAllText(path, Encoding.UTF8)) as Dictionary<string, object>;
            if (root == null) throw new InvalidOperationException("agents.json no es un objeto");
            var file = new AgentFile();
            object def;
            if (root.TryGetValue("default", out def) && def is string && ((string)def).Length > 0)
                file.DefaultId = (string)def;
            object list;
            if (!root.TryGetValue("agents", out list)) throw new InvalidOperationException("agents.json no tiene agents");
            foreach (object item in (IEnumerable)list)
            {
                var d = item as Dictionary<string, object>;
                if (d == null) continue;
                Agent a = Parse(d);
                ApplyDefaults(a);
                file.Agents.Add(a);
            }
            if (file.Agents.Count == 0) throw new InvalidOperationException("agents.json no declara agentes");
            return file;
        }

        static Agent Parse(Dictionary<string, object> d)
        {
            var a = new Agent();
            a.Id = Str(d, "id");
            a.Label = Str(d, "label");
            if (string.IsNullOrEmpty(a.Label)) a.Label = a.Id ?? "agente";
            a.Enabled = Bool(d, "enabled", true);
            a.Command = Str(d, "command") ?? "";
            object args;
            if (d.TryGetValue("args", out args) && args is IEnumerable && !(args is string))
            {
                foreach (object item in (IEnumerable)args)
                    if (item != null) a.Args.Add(Convert.ToString(item, CultureInfo.InvariantCulture));
            }
            a.PromptVia = Str(d, "promptVia") ?? a.PromptVia;
            a.PromptFormat = Str(d, "promptFormat") ?? a.PromptFormat;
            a.PromptArg = Str(d, "promptArg");
            a.Parser = Str(d, "parser") ?? a.Parser;
            a.Stream = Bool(d, "stream", true);
            a.Images = Str(d, "images") ?? a.Images;
            a.ImageArg = Str(d, "imageArg") ?? a.ImageArg;
            a.MaxImages = IntVal(d, "maxImages", 0);
            a.Model = Str(d, "model");
            a.ModelArg = Str(d, "modelArg");
            a.Rules = Str(d, "rules") ?? a.Rules;
            a.RulesArg = Str(d, "rulesArg") ?? a.RulesArg;
            a.Context = Str(d, "context") ?? a.Context;
            object env;
            if (d.TryGetValue("env", out env))
            {
                var map = env as Dictionary<string, object>;
                if (map != null)
                {
                    foreach (KeyValuePair<string, object> kv in map)
                        if (kv.Value is string) a.Env[kv.Key] = (string)kv.Value;
                }
            }
            a.TextRules = Rules(d, "textRules");
            a.ErrorRules = Rules(d, "errorRules");
            a.StatusRules = Rules(d, "statusRules");
            return a;
        }

        public static void ApplyDefaults(Agent a)
        {
            if (a.Parser == "codex-jsonl")
            {
                if (a.TextRules.Count == 0) a.TextRules.Add(Rule("item.completed", "item.text", "append", "item.type", "agent_message"));
                if (a.ErrorRules.Count == 0)
                {
                    a.ErrorRules.Add(Rule("turn.failed", "error.message", "set", null, null));
                    a.ErrorRules.Add(Rule("error", "message", "set", null, null));
                }
                if (a.StatusRules.Count == 0)
                    a.StatusRules.Add(Rule("item.started", null, null, "item.type", "command_execution"));
            }
            else if (a.Parser == "gemini-stream")
            {
                if (a.TextRules.Count == 0)
                {
                    a.TextRules.Add(Rule("message", "content", "auto", "role", "assistant"));
                    a.TextRules.Add(Rule("result", "response", "set-if-empty", "status", "success"));
                }
                if (a.ErrorRules.Count == 0)
                {
                    a.ErrorRules.Add(Rule("result", "error.message", "set", "status", "error"));
                    a.ErrorRules.Add(Rule("error", "message", "set", "severity", "error"));
                }
                if (a.StatusRules.Count == 0) a.StatusRules.Add(Rule("tool_use", null, null, null, null));
            }
            else if (a.Parser == "grok-stream")
            {
                if (a.TextRules.Count == 0) a.TextRules.Add(Rule("text", "data", "append", null, null));
                if (a.ErrorRules.Count == 0) a.ErrorRules.Add(Rule("error", "message", "set", null, null));
                if (a.StatusRules.Count == 0)
                {
                    JsonlRule s = Rule("tool_call", null, null, null, null);
                    s.Clear = true;
                    a.StatusRules.Add(s);
                }
            }
            else if (a.Parser == "hermes-stream")
            {
                if (a.TextRules.Count == 0)
                {
                    a.TextRules.Add(Rule("text", "text", "append", null, null));
                    a.TextRules.Add(Rule("result", "text", "set", null, null));
                }
                if (a.ErrorRules.Count == 0) a.ErrorRules.Add(Rule("result", "error", "set", null, null));
                if (a.StatusRules.Count == 0)
                {
                    JsonlRule s = Rule("tool_use", null, null, null, null);
                    s.Clear = true;
                    a.StatusRules.Add(s);
                }
            }
        }

        static JsonlRule Rule(string type, string path, string mode, string matchKey, string matchVal)
        {
            var r = new JsonlRule();
            r.Type = type;
            r.Path = path;
            r.Mode = mode;
            if (matchKey != null) r.Match[matchKey] = matchVal;
            return r;
        }

        static List<JsonlRule> Rules(Dictionary<string, object> d, string key)
        {
            var list = new List<JsonlRule>();
            object v;
            if (!d.TryGetValue(key, out v) || !(v is IEnumerable) || v is string) return list;
            foreach (object item in (IEnumerable)v)
            {
                var o = item as Dictionary<string, object>;
                if (o == null) continue;
                var r = new JsonlRule();
                r.Type = Str(o, "type");
                r.Path = Str(o, "path");
                r.Mode = Str(o, "mode") ?? "append";
                r.Clear = Bool(o, "clear", false);
                object match;
                if (o.TryGetValue("match", out match))
                {
                    var map = match as Dictionary<string, object>;
                    if (map != null) r.Match = map;
                }
                list.Add(r);
            }
            return list;
        }

        public Agent Resolve(string id)
        {
            if (string.IsNullOrEmpty(id)) id = DefaultId;
            if (string.IsNullOrEmpty(id)) id = "claude";
            foreach (Agent a in Agents)
            {
                if (string.Equals(a.Id, id, StringComparison.OrdinalIgnoreCase)) return a;
            }
            return null;
        }

        public string Ids()
        {
            var sb = new StringBuilder();
            foreach (Agent a in Agents)
            {
                if (sb.Length > 0) sb.Append(", ");
                sb.Append(a.Id);
            }
            return sb.ToString();
        }

        static string Str(Dictionary<string, object> o, string key)
        {
            object v;
            return o.TryGetValue(key, out v) ? v as string : null;
        }

        static bool Bool(Dictionary<string, object> o, string key, bool def)
        {
            object v;
            if (!o.TryGetValue(key, out v) || v == null) return def;
            if (v is bool) return (bool)v;
            return def;
        }

        static int IntVal(Dictionary<string, object> o, string key, int def)
        {
            object v;
            if (!o.TryGetValue(key, out v) || v == null) return def;
            try { return Convert.ToInt32(v, CultureInfo.InvariantCulture); }
            catch { return def; }
        }
    }

    class Attachment
    {
        public string MediaType;
        public string Base64;

        const int MaxSide = 1568;
        const int MaxImages = 20;
        const string Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

        public static List<Attachment> FromBookmark(dynamic doc, string name)
        {
            var list = new List<Attachment>();
            try
            {
                if (!doc.Bookmarks.Exists(name)) return list;

                dynamic r = doc.Bookmarks.Item(name).Range;
                dynamic full = doc.Range(r.Paragraphs.First.Range.Start, r.Paragraphs.Last.Range.End);
                string xml = full.WordOpenXML;

                var x = new XmlDocument();
                x.LoadXml(xml);
                var ns = new XmlNamespaceManager(x.NameTable);
                ns.AddNamespace("pkg", Pkg);

                foreach (XmlNode node in x.SelectNodes("//pkg:part[starts-with(@pkg:contentType, 'image/')]", ns))
                {
                    XmlNode data = node.SelectSingleNode("pkg:binaryData", ns);
                    if (data == null) continue;
                    string type = ((XmlElement)node).GetAttribute("contentType", Pkg);
                    Attachment a = Normalize(type, Convert.FromBase64String(data.InnerText));
                    if (a != null) list.Add(a);
                    if (list.Count >= MaxImages) break;
                }
                if (list.Count == 0) Program.Log("marcador sin imagenes en el XML");
            }
            catch (Exception ex) { Program.Log("imagenes: " + ex.Message); }
            return list;
        }

        static Attachment Normalize(string type, byte[] bytes)
        {
            try
            {
                using (var ms = new MemoryStream(bytes))
                using (var img = System.Drawing.Image.FromStream(ms))
                {
                    bool native = type == "image/png" || type == "image/jpeg" || type == "image/gif";
                    int side = Math.Max(img.Width, img.Height);
                    if (native && side <= MaxSide && bytes.Length < 3500000)
                        return new Attachment { MediaType = type, Base64 = Convert.ToBase64String(bytes) };

                    double scale = Math.Min(1.0, MaxSide / (double)side);
                    int w = Math.Max(1, (int)(img.Width * scale));
                    int h = Math.Max(1, (int)(img.Height * scale));
                    using (var bmp = new Bitmap(w, h))
                    {
                        using (Graphics g = Graphics.FromImage(bmp))
                        {
                            g.Clear(Color.White);
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.DrawImage(img, 0, 0, w, h);
                        }
                        using (var png = new MemoryStream())
                        {
                            bmp.Save(png, ImageFormat.Png);
                            return new Attachment { MediaType = "image/png", Base64 = Convert.ToBase64String(png.ToArray()) };
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Program.Log("imagen omitida (" + type + "): " + ex.Message);
                return null;
            }
        }
    }

    class Typewriter
    {
        readonly Config cfg;
        readonly Stopwatch clock = Stopwatch.StartNew();
        readonly Random rnd = new Random();
        int pos;
        double budget;
        long last;
        long pauseUntil;

        public Typewriter(Config cfg) { this.cfg = cfg; }

        public void Reset() { pos = 0; budget = 0; pauseUntil = 0; }

        public int Next(string target, bool complete)
        {
            long now = clock.ElapsedMilliseconds;
            double dt = (now - last) / 1000.0;
            last = now;
            if (pos > target.Length) pos = target.Length;
            if (now < pauseUntil) return pos;

            int backlog = target.Length - pos;
            double speed = cfg.CharsPerSecond * Math.Max(1.0, backlog / (double)cfg.CatchUpAfterChars);
            double cap = speed * 0.5;
            budget = Math.Min(budget + speed * dt, cap);

            while (pos < target.Length)
            {
                int end = pos;
                while (end < target.Length && !char.IsWhiteSpace(target[end])) end++;
                if (end == target.Length && !complete) break;
                while (end < target.Length && char.IsWhiteSpace(target[end])) end++;

                int len = end - pos;
                if (budget < len && budget < cap && pos > 0) break;
                budget -= len;
                int wordStart = pos;
                pos = end;

                int pause = PauseAfter(target, wordStart, end);
                if (pause > 0)
                {
                    pauseUntil = now + (int)(pause * (0.7 + 0.6 * rnd.NextDouble()));
                    budget = 0;
                    break;
                }
            }
            return pos;
        }

        int PauseAfter(string t, int start, int end)
        {
            if (t.IndexOf('\r', start, end - start) >= 0) return cfg.ParagraphPauseMs;
            int i = end - 1;
            while (i >= start && char.IsWhiteSpace(t[i])) i--;
            if (i < start) return 0;
            char c = t[i];
            if (c == '.' || c == '?' || c == '!' || c == ':') return cfg.SentencePauseMs;
            if (c == ',' || c == ';') return cfg.SentencePauseMs / 3;
            return 0;
        }
    }

    class Config
    {
        public string ClaudePath;
        public string Workspace;
        public string Model;
        public string ClaudeConfigDir;
        public string DefaultAgent;
        public int TimeoutSeconds = 600;
        public int CharsPerSecond = 20;
        public int StartDelayMs = 800;
        public int SentencePauseMs = 450;
        public int ParagraphPauseMs = 900;
        public int CatchUpAfterChars = 600;

        public static Config Load(string path)
        {
            var o = new JavaScriptSerializer().DeserializeObject(File.ReadAllText(path, Encoding.UTF8))
                    as Dictionary<string, object>;
            var c = new Config();
            object v;
            if (o.TryGetValue("claudePath", out v)) c.ClaudePath = v as string;
            if (o.TryGetValue("workspace", out v)) c.Workspace = v as string;
            if (o.TryGetValue("model", out v)) c.Model = v as string;
            if (o.TryGetValue("claudeConfigDir", out v)) c.ClaudeConfigDir = v as string;
            if (o.TryGetValue("defaultAgent", out v)) c.DefaultAgent = v as string;
            c.TimeoutSeconds = Int(o, "timeoutSeconds", c.TimeoutSeconds);
            c.CharsPerSecond = Math.Max(1, Int(o, "charsPerSecond", c.CharsPerSecond));
            c.StartDelayMs = Int(o, "startDelayMs", c.StartDelayMs);
            c.SentencePauseMs = Int(o, "sentencePauseMs", c.SentencePauseMs);
            c.ParagraphPauseMs = Int(o, "paragraphPauseMs", c.ParagraphPauseMs);
            c.CatchUpAfterChars = Math.Max(1, Int(o, "catchUpAfterChars", c.CatchUpAfterChars));
            return c;
        }

        static int Int(Dictionary<string, object> o, string key, int def)
        {
            object v;
            if (!o.TryGetValue(key, out v) || v == null) return def;
            try { return Convert.ToInt32(v); } catch { return def; }
        }
    }
}
