// Ventana minima de Pluma (WinForms, .NET Framework, C# 5).
// Se abre con Pluma.exe --config, desde el menu de Word o desde el menu Inicio.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Pluma
{
    static class ConfigApp
    {
        public static int Run()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try { StartMenuLink.Ensure(); } catch { }

            string workspace = InstallLayout.Workspace();
            if (string.IsNullOrEmpty(workspace) || !Directory.Exists(workspace))
            {
                MessageBox.Show(
                    "No encuentro la carpeta del workspace. Volvé a correr el instalador.",
                    "Pluma", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 1;
            }
            string agents = Path.Combine(workspace, "agents.json");
            if (!File.Exists(agents))
            {
                MessageBox.Show(
                    "En el workspace no está agents.json. Volvé a correr el instalador.",
                    "Pluma", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 1;
            }
            Application.Run(new ConfigForm(workspace));
            return 0;
        }
    }

    static class InstallLayout
    {
        public static string Dir()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Pluma");
        }

        public static string Exe()
        {
            return Path.Combine(Dir(), "Pluma.exe");
        }

        public static string Workspace()
        {
            string path = Path.Combine(Dir(), "config.json");
            if (!File.Exists(path)) return null;
            try
            {
                var o = new System.Web.Script.Serialization.JavaScriptSerializer()
                    .DeserializeObject(File.ReadAllText(path, Encoding.UTF8)) as Dictionary<string, object>;
                object v;
                if (o != null && o.TryGetValue("workspace", out v)) return v as string;
            }
            catch { }
            return null;
        }
    }

    static class StartMenuLink
    {
        public static void Ensure()
        {
            string programs = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
            if (string.IsNullOrEmpty(programs)) return;
            string exe = InstallLayout.Exe();
            if (!File.Exists(exe)) exe = Application.ExecutablePath;
            Type shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null) return;
            object shell = Activator.CreateInstance(shellType);
            object link = shellType.InvokeMember("CreateShortcut", System.Reflection.BindingFlags.InvokeMethod,
                null, shell, new object[] { Path.Combine(programs, "Pluma.lnk") });
            Type linkType = link.GetType();
            linkType.InvokeMember("TargetPath", System.Reflection.BindingFlags.SetProperty, null, link, new object[] { exe });
            linkType.InvokeMember("Arguments", System.Reflection.BindingFlags.SetProperty, null, link, new object[] { "--config" });
            linkType.InvokeMember("WorkingDirectory", System.Reflection.BindingFlags.SetProperty, null, link, new object[] { Path.GetDirectoryName(exe) });
            linkType.InvokeMember("Description", System.Reflection.BindingFlags.SetProperty, null, link, new object[] { "Configuracion de Pluma" });
            linkType.InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, null, link, new object[0]);
        }
    }

    static class WordApply
    {
        public static bool TryRun()
        {
            try
            {
                dynamic word = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                word.Run("AplicarAjustesPluma");
                return true;
            }
            catch { return false; }
        }
    }

    internal static class UiColor
    {
        public static int ToOle(Color c)
        {
            return c.R | (c.G << 8) | (c.B << 16);
        }

        public static Color FromOle(int ole)
        {
            return Color.FromArgb(ole & 255, (ole >> 8) & 255, (ole >> 16) & 255);
        }
    }

    internal static class Shortcut
    {
        public static string Format(bool ctrl, bool alt, bool shift, int key)
        {
            string name = KeyName(key);
            if (name == null) return null;
            if (!ctrl && !alt && !shift) return null;
            var sb = new StringBuilder();
            if (ctrl) sb.Append("Ctrl+");
            if (alt) sb.Append("Alt+");
            if (shift) sb.Append("Shift+");
            sb.Append(name);
            return sb.ToString();
        }

        static string KeyName(int key)
        {
            if (key >= 65 && key <= 90) return ((char)key).ToString();
            if (key >= 48 && key <= 57) return ((char)key).ToString();
            if (key >= 112 && key <= 123) return "F" + (key - 111).ToString(CultureInfo.InvariantCulture);
            if (key == 13) return "Enter";
            if (key == 9) return "Tab";
            if (key == 32) return "Espacio";
            if (key == 27) return "Esc";
            return null;
        }
    }

    internal class Settings
    {
        public string DefaultAgent = "";
        public string Shortcut = "Alt+Enter";
        public bool HasStreamColor;
        public int StreamColor;
        public bool AllowInternet = true;
        Dictionary<string, object> raw = new Dictionary<string, object>();

        public static Settings Load(string path)
        {
            var s = new Settings();
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return s;
            try
            {
                var o = new System.Web.Script.Serialization.JavaScriptSerializer()
                    .DeserializeObject(File.ReadAllText(path, Encoding.UTF8)) as Dictionary<string, object>;
                if (o == null) return s;
                s.raw = o;
                object v;
                if (o.TryGetValue("defaultAgent", out v) && v is string) s.DefaultAgent = (string)v;
                if (o.TryGetValue("shortcut", out v) && v is string && ((string)v).Length > 0) s.Shortcut = (string)v;
                if (o.TryGetValue("streamColor", out v) && v != null && !(v is string))
                {
                    s.StreamColor = Convert.ToInt32(v, CultureInfo.InvariantCulture);
                    s.HasStreamColor = true;
                }
                if (o.TryGetValue("allowInternet", out v))
                {
                    if (v is bool) s.AllowInternet = (bool)v;
                    else if (v is string && string.Equals((string)v, "false", StringComparison.OrdinalIgnoreCase)) s.AllowInternet = false;
                }
            }
            catch { }
            return s;
        }

        public void Save(string path)
        {
            if (raw == null) raw = new Dictionary<string, object>();
            raw["defaultAgent"] = DefaultAgent ?? "";
            raw["shortcut"] = string.IsNullOrEmpty(Shortcut) ? "Alt+Enter" : Shortcut;
            if (HasStreamColor) raw["streamColor"] = StreamColor;
            else raw.Remove("streamColor");
            raw["allowInternet"] = AllowInternet;

            var sb = new StringBuilder();
            sb.Append("{\r\n");
            bool first = true;
            string[] order = new[] { "defaultAgent", "shortcut", "streamColor", "allowInternet" };
            var seen = new Dictionary<string, bool>();
            foreach (string key in order) WriteKey(sb, ref first, seen, key);
            foreach (string key in new List<string>(raw.Keys)) WriteKey(sb, ref first, seen, key);
            sb.Append("\r\n}\r\n");
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
        }

        void WriteKey(StringBuilder sb, ref bool first, Dictionary<string, bool> seen, string key)
        {
            object value;
            if (seen.ContainsKey(key) || !raw.TryGetValue(key, out value)) return;
            seen[key] = true;
            if (!first) sb.Append(",\r\n");
            first = false;
            sb.Append("  ").Append(JsonString(key)).Append(": ").Append(JsonValue(value));
        }

        static string JsonString(string s)
        {
            if (s == null) s = "";
            var sb = new StringBuilder();
            sb.Append('"');
            foreach (char c in s)
            {
                if (c == '\\' || c == '"') { sb.Append('\\'); sb.Append(c); }
                else if (c == '\n') sb.Append("\\n");
                else if (c == '\r') sb.Append("\\r");
                else sb.Append(c);
            }
            sb.Append('"');
            return sb.ToString();
        }

        static string JsonValue(object value)
        {
            if (value == null) return "null";
            if (value is bool) return (bool)value ? "true" : "false";
            if (value is string) return JsonString((string)value);
            if (value is int || value is long || value is decimal || value is double || value is float)
                return Convert.ToString(value, CultureInfo.InvariantCulture);
            return new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(value);
        }
    }

    class AgentPick
    {
        public string Id;
        public string Caption;
        public override string ToString() { return Caption; }
    }

    class ConfigForm : Form
    {
        readonly string settingsPath;
        readonly Settings settings;
        readonly ComboBox agents = new ComboBox();
        readonly TextBox shortcut = new TextBox();
        readonly CheckBox ownColor = new CheckBox();
        readonly CheckBox internet = new CheckBox();
        readonly Button swatch = new Button();
        readonly Label status = new Label();

        public ConfigForm(string workspace)
        {
            settingsPath = Path.Combine(workspace, "settings.json");
            settings = Settings.Load(settingsPath);

            Text = "Pluma";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(440, 372);
            Font = new Font("Segoe UI", 9f);

            var title = new Label();
            title.Text = "Configuración";
            title.Font = new Font(Font, FontStyle.Bold);
            title.SetBounds(16, 14, 400, 20);

            var agentLabel = new Label();
            agentLabel.Text = "Agente por defecto";
            agentLabel.SetBounds(16, 44, 400, 18);

            agents.DropDownStyle = ComboBoxStyle.DropDownList;
            agents.SetBounds(16, 64, 408, 24);
            FillAgents(workspace);

            internet.Text = "Permitir internet";
            internet.Checked = settings.AllowInternet;
            internet.SetBounds(16, 96, 400, 22);

            var netHint = new Label();
            netHint.Text = "Búsqueda y lectura de páginas. No habilita editar archivos ni correr comandos.";
            netHint.ForeColor = Color.DimGray;
            netHint.SetBounds(16, 118, 408, 32);

            var keyLabel = new Label();
            keyLabel.Text = "Atajo";
            keyLabel.SetBounds(16, 154, 400, 18);

            shortcut.ReadOnly = true;
            shortcut.ShortcutsEnabled = false;
            shortcut.Text = string.IsNullOrEmpty(settings.Shortcut) ? "Alt+Enter" : settings.Shortcut;
            shortcut.SetBounds(16, 174, 200, 24);
            shortcut.KeyDown += OnShortcut;

            var keyHint = new Label();
            keyHint.Text = "Hacé clic y presioná la combinación. De fábrica es Alt+Enter.";
            keyHint.ForeColor = Color.DimGray;
            keyHint.SetBounds(16, 200, 408, 18);

            ownColor.Text = "Color propio mientras se escribe la respuesta";
            ownColor.SetBounds(16, 228, 400, 22);
            ownColor.Checked = settings.HasStreamColor;

            swatch.SetBounds(16, 254, 72, 26);
            swatch.FlatStyle = FlatStyle.Flat;
            swatch.BackColor = settings.HasStreamColor ? UiColor.FromOle(settings.StreamColor) : Color.FromArgb(47, 84, 150);
            swatch.Text = "";
            swatch.Click += OnPickColor;

            var colorHint = new Label();
            colorHint.Text = "Al terminar, el texto vuelve al color del párrafo.";
            colorHint.ForeColor = Color.DimGray;
            colorHint.SetBounds(96, 258, 320, 18);

            var save = new Button();
            save.Text = "Guardar";
            save.SetBounds(16, 296, 100, 28);
            save.Click += OnSave;

            status.SetBounds(124, 300, 300, 22);
            status.ForeColor = Color.DimGray;

            var credit = new LinkLabel();
            credit.Text = "Hecho por @gabogabucho";
            credit.LinkBehavior = LinkBehavior.HoverUnderline;
            credit.SetBounds(16, 340, 400, 20);
            credit.Links.Add(10, 12, "https://x.com/gabogabucho");
            credit.LinkClicked += OnCredit;

            Controls.Add(title);
            Controls.Add(agentLabel);
            Controls.Add(agents);
            Controls.Add(internet);
            Controls.Add(netHint);
            Controls.Add(keyLabel);
            Controls.Add(shortcut);
            Controls.Add(keyHint);
            Controls.Add(ownColor);
            Controls.Add(swatch);
            Controls.Add(colorHint);
            Controls.Add(save);
            Controls.Add(status);
            Controls.Add(credit);

            AcceptButton = save;
        }

        void FillAgents(string workspace)
        {
            var cfg = new Config();
            cfg.Workspace = workspace;
            AgentFile file = AgentFile.Load(Path.Combine(workspace, "agents.json"), cfg);
            AgentPick selected = null;
            string want = settings.DefaultAgent;
            if (string.IsNullOrEmpty(want)) want = file.DefaultId;
            foreach (Agent agent in file.Agents)
            {
                if (!agent.Enabled || string.IsNullOrEmpty(agent.Command)) continue;
                var pick = new AgentPick();
                pick.Id = agent.Id;
                pick.Caption = agent.Label + "  (" + agent.Id + ")";
                agents.Items.Add(pick);
                if (string.Equals(agent.Id, want, StringComparison.OrdinalIgnoreCase)) selected = pick;
            }
            if (selected != null) agents.SelectedItem = selected;
            else if (agents.Items.Count > 0) agents.SelectedIndex = 0;
            if (agents.Items.Count == 0)
                status.Text = "No hay agentes habilitados en agents.json.";
        }

        void OnShortcut(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Menu || e.KeyCode == Keys.ControlKey || e.KeyCode == Keys.ShiftKey
                || e.KeyCode == Keys.LWin || e.KeyCode == Keys.RWin)
            {
                e.SuppressKeyPress = true;
                return;
            }
            string text = Shortcut.Format(e.Control, e.Alt, e.Shift, (int)e.KeyCode);
            if (text == null)
                status.Text = "Usá Alt, Ctrl o Shift más una tecla (Enter, una letra o F1…F12).";
            else
            {
                shortcut.Text = text;
                status.Text = "";
            }
            e.SuppressKeyPress = true;
            e.Handled = true;
        }

        void OnPickColor(object sender, EventArgs e)
        {
            using (var dialog = new ColorDialog())
            {
                dialog.Color = swatch.BackColor;
                dialog.FullOpen = true;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    swatch.BackColor = dialog.Color;
                    ownColor.Checked = true;
                }
            }
        }

        void OnSave(object sender, EventArgs e)
        {
            AgentPick pick = agents.SelectedItem as AgentPick;
            if (pick == null)
            {
                status.Text = "Elegí un agente.";
                return;
            }
            if (string.IsNullOrEmpty(shortcut.Text))
            {
                status.Text = "Falta el atajo.";
                return;
            }
            settings.DefaultAgent = pick.Id;
            settings.Shortcut = shortcut.Text;
            settings.AllowInternet = internet.Checked;
            settings.HasStreamColor = ownColor.Checked;
            settings.StreamColor = UiColor.ToOle(swatch.BackColor);
            try { settings.Save(settingsPath); }
            catch (Exception ex)
            {
                status.Text = "No pude guardar: " + ex.Message;
                return;
            }
            status.Text = WordApply.TryRun()
                ? "Guardado. El atajo ya está en Word."
                : "Guardado. El atajo se aplica al abrir Word.";
        }

        void OnCredit(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                string url = e.Link.LinkData as string;
                if (!string.IsNullOrEmpty(url)) Process.Start(url);
            }
            catch { }
        }
    }
}
