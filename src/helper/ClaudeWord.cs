// ClaudeWord.exe — helper oculto que ejecuta Claude Code y escribe la
// respuesta en stream dentro de Word via COM.
//
// Lo lanza la macro EnviarAClaude con un id. Con ese id encuentra el bloque
// (content control con tag "claude-<id>") y las variables del documento con
// el texto y la instruccion. Compila con el csc.exe de .NET Framework 4
// (C# 5: sin interpolacion, sin ?., sin nameof).

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;

namespace ClaudeWord
{
    static class Program
    {
        const string TagPrefix = "claude-";
        const int TickMs = 40;
        const int ColorStatus = 15263976;       // #E8E8E8, casi blanco: avisos de estado
        const int ColorAutomatic = -16777216;   // wdColorAutomatic: la respuesta va en negro
        const string StatusReading = "✳ Claude está leyendo el contexto…";

        static readonly object Gate = new object();
        static readonly StringBuilder Answer = new StringBuilder();
        static readonly StringBuilder StdErr = new StringBuilder();
        static string status;
        static string resultText;
        static string error;
        static bool finished;
        static string logPath;

        [STAThread]
        static int Main(string[] args)
        {
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
                if (doc != null) Finish(doc, tag, "⚠ Claude no pudo responder: " + ex.Message);
                return 1;
            }
        }

        static int Run(dynamic doc, string tag, string baseDir)
        {
            Config cfg = Config.Load(Path.Combine(baseDir, "config.json"));
            string texto = ReadVariable(doc, tag + "-text") ?? "";
            string instr = ReadVariable(doc, tag + "-instr");

            if (!File.Exists(cfg.ClaudePath))
            {
                Finish(doc, tag, "⚠ No encontré Claude Code en esta PC (" + cfg.ClaudePath + ").");
                return 4;
            }

            Process p = StartClaude(cfg, baseDir, BuildPrompt(instr, texto));
            Stopwatch clock = Stopwatch.StartNew();
            var typer = new Typewriter(cfg);
            string shown = null;          // lo que esta escrito ahora en el bloque
            bool showingStatus = false;
            long firstTextAt = -1;        // cuando llego la primera palabra
            string target = "";

            // La respuesta llega de Claude a su ritmo; el Typewriter decide cuanto
            // se muestra. Se termina cuando Claude termino Y se mostro todo.
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
                            // Solo texto nuevo al final: se agrega. Si cambio algo anterior
                            // (markdown limpiado, estado reemplazado) se reescribe.
                            if (shown != null && !showingStatus && visible.StartsWith(shown, StringComparison.Ordinal))
                                AppendBlock(cc, visible.Substring(shown.Length));
                            else
                                WriteBlock(cc, visible, ColorAutomatic);
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

            p.WaitForExit(5000);

            string err;
            lock (Gate)
            {
                err = error;
                if (err == null && target.Length == 0 && StdErr.Length > 0) err = StdErr.ToString().Trim();
            }
            Log("fin. error=" + (err ?? "-") + " chars=" + target.Length);

            if (err != null && target.Length == 0)
                Finish(doc, tag, "⚠ Claude no pudo responder: " + err);
            else
                Finish(doc, tag, target == shown ? null : target);
            return 0;
        }

        // ---------- Claude Code ----------

        static string BuildPrompt(string instr, string texto)
        {
            // Sin instruccion, lo que ella escribio ES el mensaje.
            if (string.IsNullOrEmpty(instr)) return texto.Trim();
            return "Instrucción: " + instr + "\n\n<texto>\n" + texto.Trim() + "\n</texto>";
        }

        static Process StartClaude(Config cfg, string baseDir, string prompt)
        {
            var args = new StringBuilder();
            args.Append("-p --output-format stream-json --include-partial-messages --verbose");
            args.Append(" --setting-sources project --strict-mcp-config --no-session-persistence");
            // Solo lectura + skills del workspace (.claude/skills). Nada de editar ni ejecutar.
            args.Append(" --tools \"Read,Glob,Grep,Skill\"");
            args.Append(" --append-system-prompt-file ").Append(Quote(Path.Combine(baseDir, "word-rules.md")));
            if (!string.IsNullOrEmpty(cfg.Model)) args.Append(" --model ").Append(Quote(cfg.Model));

            var psi = new ProcessStartInfo(cfg.ClaudePath, args.ToString());
            psi.WorkingDirectory = cfg.Workspace;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.RedirectStandardInput = true;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;
            psi.StandardOutputEncoding = Encoding.UTF8;
            psi.StandardErrorEncoding = Encoding.UTF8;
            if (!string.IsNullOrEmpty(cfg.ClaudeConfigDir))
                psi.EnvironmentVariables["CLAUDE_CONFIG_DIR"] = cfg.ClaudeConfigDir;

            Log("claude " + args);
            Process p = Process.Start(psi);

            p.ErrorDataReceived += (s, e) => { if (e.Data != null) lock (Gate) StdErr.AppendLine(e.Data); };
            p.BeginErrorReadLine();

            var reader = new Thread(() =>
            {
                var json = new JavaScriptSerializer();
                json.MaxJsonLength = int.MaxValue;
                string line;
                while ((line = p.StandardOutput.ReadLine()) != null) HandleEvent(json, line);
                lock (Gate) finished = true;
            });
            reader.IsBackground = true;
            reader.Start();

            // .NET Framework no permite elegir el encoding de stdin: se escribe en bytes UTF-8.
            byte[] bytes = new UTF8Encoding(false).GetBytes(prompt);
            p.StandardInput.BaseStream.Write(bytes, 0, bytes.Length);
            p.StandardInput.Close();
            return p;
        }

        // Eventos de "claude -p --output-format stream-json --include-partial-messages".
        // El texto que Claude escribe antes de usar una herramienta ("voy a leer...")
        // se descarta: solo queda el texto posterior a la ultima herramienta.
        static void HandleEvent(JavaScriptSerializer json, string line)
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
                        lock (Gate) { Answer.Clear(); status = StatusReading; }
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

        // Word crea un parrafo por cada salto de linea: se colapsan los vacios
        // y se limpia el markdown que se haya colado.
        static string Clean(string s)
        {
            s = s.Replace("\r\n", "\n");
            s = Regex.Replace(s, @"(?m)^#{1,6}\s+", "");
            s = Regex.Replace(s, @"\*\*(.+?)\*\*", "$1");
            s = Regex.Replace(s, @"\n{2,}", "\n");
            return s.Trim().Replace("\n", "\r");
        }

        static void Kill(Process p)
        {
            try { if (!p.HasExited) p.Kill(); } catch { }
        }

        static string Quote(string s) { return "\"" + s + "\""; }

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

        // Escribe el texto final (null = ya esta escrito) y quita el bloque
        // dejando el contenido como texto normal.
        static void Finish(dynamic doc, string tag, string text)
        {
            try
            {
                dynamic cc = GetBlock(doc, tag);
                if (cc != null)
                {
                    if (text != null) WriteBlock(cc, text, ColorAutomatic);
                    cc.LockContents = false;
                    cc.Delete(false);
                }
            }
            catch (Exception ex) { Log("finish: " + ex.Message); }
            DeleteVariables(doc, tag);
            try { doc.Application.StatusBar = " "; } catch { }
        }

        static string ReadVariable(dynamic doc, string name)
        {
            try { return (string)doc.Variables.Item(name).Value; }
            catch { return null; }
        }

        static void DeleteVariables(dynamic doc, string tag)
        {
            foreach (string suffix in new[] { "-text", "-instr" })
            {
                try { doc.Variables.Item(tag + suffix).Delete(); } catch { }
            }
        }

        // ---------- utilidades ----------

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

        static void Log(string msg)
        {
            try { File.AppendAllText(logPath, DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg + "\r\n"); }
            catch { }
        }

        // ---------- COM: reintentar cuando Word esta ocupado ----------
        // Si ella esta escribiendo o tiene un menu abierto, Word rechaza la llamada
        // (SERVERCALL_RETRYLATER). Este filtro hace que COM reintente solo.

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

    // Muestra la respuesta como alguien escribiendo: palabra por palabra, a un ritmo
    // fijo, con pausas despues de comas, puntos y parrafos. Si el atraso crece mucho
    // (respuestas largas) acelera en proporcion para no eternizarse.
    // Los tokens ya llegaron y se pagaron: esto solo cambia como se muestran.
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

        // Devuelve cuantos caracteres de target deben estar visibles ahora.
        // complete = Claude termino, la ultima palabra ya no va a cambiar.
        public int Next(string target, bool complete)
        {
            long now = clock.ElapsedMilliseconds;
            double dt = (now - last) / 1000.0;
            last = now;
            if (pos > target.Length) pos = target.Length;
            if (now < pauseUntil) return pos;

            int backlog = target.Length - pos;
            double speed = cfg.CharsPerSecond * Math.Max(1.0, backlog / (double)cfg.CatchUpAfterChars);
            double cap = speed * 0.5;                    // sin rafagas tras un COM lento
            budget = Math.Min(budget + speed * dt, cap);

            while (pos < target.Length)
            {
                int end = pos;
                while (end < target.Length && !char.IsWhiteSpace(target[end])) end++;
                if (end == target.Length && !complete) break;       // palabra todavia llegando
                while (end < target.Length && char.IsWhiteSpace(target[end])) end++;

                // Una palabra mas larga que el tope sale apenas el presupuesto se llena.
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
