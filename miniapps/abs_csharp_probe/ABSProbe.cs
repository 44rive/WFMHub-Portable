// WFMHub: read-only diagnostic using the installed .NET Framework, not Python.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Automation;

[assembly: AssemblyVersion("0.1.0.0")]
[assembly: AssemblyTitle("WFMHub ABS CSharp Probe")]

namespace WfmHubAbsProbe
{
    public static class Json
    {
        public static JavaScriptSerializer Serializer()
        {
            return new JavaScriptSerializer { MaxJsonLength = 8 * 1024 * 1024, RecursionLimit = 80 };
        }
        public static string Encode(object value) { return Serializer().Serialize(value); }
        public static Dictionary<string, object> Read(string path)
        {
            if (new FileInfo(path).Length > 8 * 1024 * 1024) throw new InvalidDataException();
            return Serializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(path, Encoding.UTF8));
        }
        public static void Write(string path, object value)
        {
            string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temp, Encode(value) + Environment.NewLine, new UTF8Encoding(false));
                File.Move(temp, path);
            }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }
    }

    public static class ProbeLogic
    {
        private static readonly Dictionary<string, string> Headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            {"Task Group", "Task Group"}, {"Task type", "Task type"}, {"Task priority", "Task priority"},
            {"Language", "Language"}, {"Scheduled for", "Scheduled for"}, {"Assigned to", "Assigned to"},
            {"Case number", "Case number"}, {"Task description", "Task description"}, {"Contract Name", "Contract Name"}
        };
        public static Dictionary<string, object> Describe(object value)
        {
            string text = value as string;
            if (text == null) return new Dictionary<string, object> { {"present", false}, {"length", 0} };
            string header;
            Headers.TryGetValue(text.Trim(), out header);
            return new Dictionary<string, object>
            {
                {"present", text.Length > 0}, {"length", text.Length}, {"header", header},
                {"redacted", text.Length > 0 && header == null},
                {"ends_in_ellipsis", text.TrimEnd().EndsWith("...", StringComparison.Ordinal) || text.TrimEnd().EndsWith("\u2026", StringComparison.Ordinal)}
            };
        }
        public static string Error(Exception error)
        {
            // Provider messages can contain private control text: persist type/code only.
            return error.GetType().Name + " / 0x" + error.HResult.ToString("X8", CultureInfo.InvariantCulture);
        }
    }

    public static class AccessibilityProbe
    {
        private sealed class Pending
        {
            public AutomationElement Element;
            public int? Parent;
            public int Depth;
        }
        private static object Property(AutomationElement element, AutomationProperty property)
        {
            object value = element.GetCurrentPropertyValue(property, true);
            return value == AutomationElement.NotSupported ? null : value;
        }
        private static Dictionary<string, object> ReadNode(AutomationElement element)
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            bool password = Object.Equals(Property(element, AutomationElement.IsPasswordProperty), true);
            result["password_control"] = password;
            if (password) return result;
            ControlType kind = Property(element, AutomationElement.ControlTypeProperty) as ControlType;
            result["control_type"] = kind == null ? null : (object)kind.Id;
            result["name"] = ProbeLogic.Describe(Property(element, AutomationElement.NameProperty));
            result["automation_id"] = ProbeLogic.Describe(Property(element, AutomationElement.AutomationIdProperty));
            result["class_name"] = ProbeLogic.Describe(Property(element, AutomationElement.ClassNameProperty));
            result["framework"] = ProbeLogic.Describe(Property(element, AutomationElement.FrameworkIdProperty));
            result["offscreen"] = Property(element, AutomationElement.IsOffscreenProperty);
            object pattern;
            bool grid = element.TryGetCurrentPattern(GridPattern.Pattern, out pattern);
            result["grid"] = grid;
            if (grid)
            {
                GridPattern.GridPatternInformation info = ((GridPattern)pattern).Current;
                result["grid_rows"] = info.RowCount;
                result["grid_columns"] = info.ColumnCount;
            }
            result["table"] = element.TryGetCurrentPattern(TablePattern.Pattern, out pattern);
            bool item = element.TryGetCurrentPattern(GridItemPattern.Pattern, out pattern);
            result["grid_item"] = item;
            if (item)
            {
                GridItemPattern.GridItemPatternInformation info = ((GridItemPattern)pattern).Current;
                result["row"] = info.Row;
                result["column"] = info.Column;
            }
            result["text_pattern"] = element.TryGetCurrentPattern(TextPattern.Pattern, out pattern);
            bool hasValue = element.TryGetCurrentPattern(ValuePattern.Pattern, out pattern);
            result["value_pattern"] = hasValue;
            if (hasValue) result["value"] = ProbeLogic.Describe(((ValuePattern)pattern).Current.Value);
            return result;
        }

        public static Dictionary<string, object> Inspect(AutomationElement root, int maxNodes, int maxDepth)
        {
            List<Dictionary<string, object>> nodes = new List<Dictionary<string, object>>();
            Stack<Pending> pending = new Stack<Pending>();
            pending.Push(new Pending { Element = root, Parent = null, Depth = 0 });
            bool limited = false;
            int readErrors = 0, navigationErrors = 0;
            TreeWalker walker = TreeWalker.RawViewWalker;
            while (pending.Count > 0 && nodes.Count < maxNodes)
            {
                Pending current = pending.Pop();
                int index = nodes.Count;
                Dictionary<string, object> node;
                try { node = ReadNode(current.Element); }
                catch (Exception error) { node = new Dictionary<string, object> { {"read_error", ProbeLogic.Error(error)} }; readErrors++; }
                node["node"] = index;
                node["parent"] = current.Parent;
                node["depth"] = current.Depth;
                nodes.Add(node);
                try
                {
                    if (current.Depth > 0)
                    {
                        AutomationElement sibling = walker.GetNextSibling(current.Element);
                        if (sibling != null) pending.Push(new Pending { Element = sibling, Parent = current.Parent, Depth = current.Depth });
                    }
                    bool password = node.ContainsKey("password_control") && Object.Equals(node["password_control"], true);
                    // An unreadable password flag must not lead into an unknown subtree.
                    if (current.Depth < maxDepth && !password && !node.ContainsKey("read_error"))
                    {
                        AutomationElement child = walker.GetFirstChild(current.Element);
                        if (child != null) pending.Push(new Pending { Element = child, Parent = index, Depth = current.Depth + 1 });
                    }
                    else if (current.Depth >= maxDepth) limited = true;
                }
                catch (Exception error) { node["navigation_error"] = ProbeLogic.Error(error); navigationErrors++; }
            }
            limited = limited || pending.Count > 0;
            SortedSet<string> headers = new SortedSet<string>(StringComparer.Ordinal);
            List<int> grids = new List<int>();
            int positioned = 0;
            foreach (Dictionary<string, object> node in nodes)
            {
                foreach (string key in new string[] {"name", "value"})
                {
                    if (!node.ContainsKey(key)) continue;
                    Dictionary<string, object> value = node[key] as Dictionary<string, object>;
                    if (value != null && value.ContainsKey("header") && value["header"] is string) headers.Add((string)value["header"]);
                }
                if (node.ContainsKey("grid") && Object.Equals(node["grid"], true)) grids.Add((int)node["node"]);
                if (node.ContainsKey("grid_item") && Object.Equals(node["grid_item"], true)) positioned++;
            }
            return new Dictionary<string, object>
            {
                {"nodes", nodes},
                {"summary", new Dictionary<string, object>
                    {
                        {"nodes_scanned", nodes.Count}, {"limit_reached", limited}, {"headers_found", new List<string>(headers)},
                        {"grid_candidates", grids}, {"positioned_cells", positioned}, {"read_errors", readErrors},
                        {"navigation_errors", navigationErrors}, {"full_backlog_coverage", "not_verified"}
                    }
                }
            };
        }

        public static Dictionary<string, object> Run(string title, int maxNodes, int maxDepth)
        {
            AutomationElementCollection windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
            List<AutomationElement> matching = new List<AutomationElement>();
            foreach (AutomationElement element in windows)
            {
                try
                {
                    string name = Property(element, AutomationElement.NameProperty) as string;
                    if (name == null || name.IndexOf(title, StringComparison.OrdinalIgnoreCase) < 0) continue;
                    object processId = Property(element, AutomationElement.ProcessIdProperty);
                    if (!(processId is int)) continue;
                    using (Process process = Process.GetProcessById((int)processId))
                    {
                        if (String.Equals(process.ProcessName, "msedge", StringComparison.OrdinalIgnoreCase)) matching.Add(element);
                    }
                }
                catch (Exception) { /* A closed/inaccessible window is not selected. */ }
            }
            if (matching.Count != 1)
            {
                return new Dictionary<string, object>
                {
                    {"status", matching.Count == 0 ? "no_matching_window" : "ambiguous_windows"}, {"matching_windows", matching.Count},
                    {"message", "Open ABS as the active tab in one visible Edge window. Close duplicate ABS windows and retry."}
                };
            }
            Dictionary<string, object> report = Inspect(matching[0], maxNodes, maxDepth);
            report["status"] = "probe_completed";
            return report;
        }
    }

    public static class Program
    {
        private static string DefaultOutput()
        {
            return Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "output", "abs-csharp-probe"));
        }
        private static int Worker(string requestPath)
        {
            Dictionary<string, object> request = Json.Read(requestPath);
            string title = Convert.ToString(request["title"], CultureInfo.InvariantCulture);
            int maxNodes = Convert.ToInt32(request["max_nodes"], CultureInfo.InvariantCulture);
            int maxDepth = Convert.ToInt32(request["max_depth"], CultureInfo.InvariantCulture);
            if (String.IsNullOrWhiteSpace(title) || title.Length > 200 || maxNodes < 1 || maxNodes > 5000 || maxDepth < 1 || maxDepth > 40) return 2;
            Dictionary<string, object> report;
            try { report = AccessibilityProbe.Run(title, maxNodes, maxDepth); }
            catch (Exception error)
            {
                report = new Dictionary<string, object> { {"status", "accessibility_unavailable"}, {"error", ProbeLogic.Error(error)},
                    {"message", "Windows accessibility could not be read. Check the error code and your workstation's application permissions."} };
            }
            Json.Write(Path.Combine(Path.GetDirectoryName(requestPath), "result.json"), report);
            return 0;
        }
        private static Dictionary<string, object> Collect(string title, int maxNodes, int maxDepth, int timeout, bool testTimeout)
        {
            string temporary = Path.Combine(Path.GetTempPath(), "wfmhub-csharp-probe-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temporary);
            try
            {
                string request = Path.Combine(temporary, "request.json");
                Json.Write(request, new Dictionary<string, object> { {"title", title}, {"max_nodes", maxNodes}, {"max_depth", maxDepth} });
                ProcessStartInfo info = new ProcessStartInfo
                {
                    FileName = Assembly.GetExecutingAssembly().Location,
                    Arguments = (testTimeout ? "--test-stalled-worker " : "--worker ") + "\"" + request + "\"",
                    UseShellExecute = false, CreateNoWindow = true
                };
                using (Process child = Process.Start(info))
                {
                    if (child == null) throw new InvalidOperationException();
                    if (!child.WaitForExit(timeout * 1000))
                    {
                        child.Kill();
                        child.WaitForExit(5000);
                        return new Dictionary<string, object> { {"status", "timeout"},
                            {"message", "The accessibility worker exceeded its time limit. Only that worker was stopped; ABS and Edge were not stopped."} };
                    }
                    string result = Path.Combine(temporary, "result.json");
                    if (child.ExitCode != 0 || !File.Exists(result))
                        return new Dictionary<string, object> { {"status", "worker_failed"}, {"worker_exit_code", child.ExitCode},
                            {"message", "The worker could not complete. Windows may have blocked a component; no ABS task capture was made."} };
                    return Json.Read(result);
                }
            }
            finally
            {
                try { Directory.Delete(temporary, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }
        private static void Save(Dictionary<string, object> report, string output)
        {
            report["schema"] = "wfmhub_abs_csharp_probe_v1";
            report["version"] = "0.1.0";
            report["captured_at"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
            report["values_redacted"] = true;
            report["runtime"] = ".NET Framework CLR " + Environment.Version;
            report["process_64_bit"] = Environment.Is64BitProcess;
            report["full_backlog_coverage"] = "not_verified";
            Directory.CreateDirectory(output);
            string path = Path.Combine(output, "probe-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Json.Write(path + ".json", report);
            StringBuilder text = new StringBuilder();
            text.AppendLine("WFMHub C# ABS diagnostic");
            text.AppendLine("Status: " + report["status"]);
            text.AppendLine("Runtime: " + report["runtime"]);
            if (report.ContainsKey("message")) text.AppendLine(Convert.ToString(report["message"], CultureInfo.InvariantCulture));
            if (report.ContainsKey("error")) text.AppendLine("Error: " + report["error"]);
            if (report.ContainsKey("summary")) text.AppendLine("Summary: " + Json.Encode(report["summary"]));
            text.AppendLine("Case text, URLs and arbitrary control values are redacted. Password controls are skipped.");
            text.AppendLine("Full backlog coverage: NOT VERIFIED. This diagnostic is not a task capture.");
            text.AppendLine("Report: " + path + ".json");
            File.WriteAllText(path + ".txt", text.ToString(), new UTF8Encoding(false));
            Console.WriteLine(text.ToString());
        }

        [MTAThread]
        public static int Main(string[] args)
        {
            try
            {
                if (args.Length == 2 && args[0] == "--worker") return Worker(args[1]);
                if (args.Length == 2 && args[0] == "--test-stalled-worker") { Thread.Sleep(120000); return 0; }
                string title = "ABS@ACM", output = DefaultOutput();
                int maxNodes = 1200, maxDepth = 20, timeout = 45;
                bool launch = false, timeoutTest = false;
                for (int i = 0; i < args.Length; i++)
                {
                    string arg = args[i];
                    if (arg == "--help") { Console.WriteLine("ABSProbe.exe [--check-launch] [--title marker] [--output folder] [--max-nodes 1200] [--max-depth 20] [--timeout 45]"); return 0; }
                    if (arg == "--check-launch") { launch = true; continue; }
                    if (arg == "--test-timeout") { timeoutTest = true; continue; }
                    if (i + 1 >= args.Length) throw new ArgumentException("Missing option value.");
                    string value = args[++i];
                    switch (arg)
                    {
                        case "--title": title = value; break;
                        case "--output": output = Path.GetFullPath(value); break;
                        case "--max-nodes": maxNodes = Int32.Parse(value, CultureInfo.InvariantCulture); break;
                        case "--max-depth": maxDepth = Int32.Parse(value, CultureInfo.InvariantCulture); break;
                        case "--timeout": timeout = Int32.Parse(value, CultureInfo.InvariantCulture); break;
                        default: throw new ArgumentException("Unknown option.");
                    }
                }
                if (String.IsNullOrWhiteSpace(title) || title.Length > 200 || maxNodes < 1 || maxNodes > 5000 || maxDepth < 1 || maxDepth > 40 || timeout < 5 || timeout > 120)
                    throw new ArgumentException("Use a title marker, 1-5000 nodes, 1-40 depth, and 5-120 seconds.");
                Stopwatch watch = Stopwatch.StartNew();
                Dictionary<string, object> report = launch
                    ? new Dictionary<string, object> { {"status", "launch_ok"}, {"message", "The C# executable and report writer launched. ABS accessibility has not been inspected."} }
                    : Collect(title, maxNodes, maxDepth, timeout, timeoutTest);
                report["elapsed_seconds"] = Math.Round(watch.Elapsed.TotalSeconds, 2);
                Save(report, output);
                string status = Convert.ToString(report["status"], CultureInfo.InvariantCulture);
                return status == "launch_ok" || status == "probe_completed" ? 0 : 1;
            }
            catch (Exception error)
            {
                Console.Error.WriteLine("The C# diagnostic did not complete: " + ProbeLogic.Error(error));
                Console.Error.WriteLine("No successful diagnostic is claimed. Check the error above and workstation application permissions.");
                return 2;
            }
        }
    }
}
