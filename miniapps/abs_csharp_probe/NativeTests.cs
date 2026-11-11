// CI only: a synthetic owned window, never the user's ABS session.
using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace WfmHubAbsProbe
{
    public static class NativeTests
    {
        private static void Assert(bool value, string message)
        {
            if (!value) throw new Exception(message);
        }
        [MTAThread]
        public static int Main()
        {
            Dictionary<string, object> text = ProbeLogic.Describe("PRIVATE-CUSTOMER-CASE-001...");
            Assert(Object.Equals(text["redacted"], true), "Text was not redacted");
            Assert(Object.Equals(text["ends_in_ellipsis"], true), "Clipping metadata missing");
            Assert(!Json.Encode(text).Contains("PRIVATE"), "Private value leaked");
            Assert(Object.Equals(ProbeLogic.Describe(" Scheduled for ")["header"], "Scheduled for"), "Header normalization failed");
            Exception sensitive = new Exception("PRIVATE-PROVIDER-CONTENT");
            Assert(!ProbeLogic.Error(sensitive).Contains("PRIVATE"), "Provider exception leaked text");
            Form window = null;
            IntPtr handle = IntPtr.Zero;
            Exception creationError = null;
            using (ManualResetEvent ready = new ManualResetEvent(false))
            {
                Thread owner = new Thread(delegate()
                {
                    try
                    {
                        window = new Form { Text = "Synthetic WFMHub diagnostic test", Width = 400, Height = 220, ShowInTaskbar = false };
                        window.Controls.Add(new Label { Text = "Task priority", AccessibleName = "Task priority", Left = 10, Top = 10, Width = 170 });
                        window.Controls.Add(new Label { Text = "PRIVATE-CASE-001", AccessibleName = "PRIVATE-CASE-001", Left = 10, Top = 45, Width = 170 });
                        window.Controls.Add(new TextBox { Text = "PRIVATE-PASSWORD-002", UseSystemPasswordChar = true, Left = 10, Top = 80, Width = 170 });
                        window.Shown += delegate { handle = window.Handle; ready.Set(); };
                        Application.Run(window);
                    }
                    catch (Exception error) { creationError = error; ready.Set(); }
                });
                owner.SetApartmentState(ApartmentState.STA);
                owner.IsBackground = true;
                owner.Start();
                try
                {
                    Assert(ready.WaitOne(10000), "Synthetic window startup timed out");
                    Assert(creationError == null, "Synthetic window failed to start");
                    Assert(handle != IntPtr.Zero, "Synthetic window has no handle");
                    AutomationElement root = AutomationElement.FromHandle(handle);
                    Dictionary<string, object> report = AccessibilityProbe.Inspect(root, 300, 15);
                    string encoded = Json.Encode(report);
                    Assert(!encoded.Contains("PRIVATE"), "Synthetic case/password data leaked");
                    Dictionary<string, object> summary = (Dictionary<string, object>)report["summary"];
                    Assert(((List<string>)summary["headers_found"]).Contains("Task priority"), "Actual Windows UIA did not expose the synthetic header");
                    Assert(Object.Equals(summary["full_backlog_coverage"], "not_verified"), "Diagnostic claimed task coverage");
                    Dictionary<string, object> bounded = AccessibilityProbe.Inspect(root, 1, 15);
                    Dictionary<string, object> boundSummary = (Dictionary<string, object>)bounded["summary"];
                    Assert(Object.Equals(boundSummary["nodes_scanned"], 1), "Node bound failed");
                    Assert(Object.Equals(boundSummary["limit_reached"], true), "Incomplete traversal was not marked");
                    bool passwordSkipped = false;
                    foreach (Dictionary<string, object> node in (List<Dictionary<string, object>>)report["nodes"])
                        if (node.ContainsKey("password_control") && Object.Equals(node["password_control"], true))
                        {
                            Assert(!node.ContainsKey("name") && !node.ContainsKey("value"), "Password control was read");
                            passwordSkipped = true;
                        }
                    Assert(passwordSkipped, "Synthetic password control was not identified");
                }
                finally
                {
                    if (window != null && handle != IntPtr.Zero) window.BeginInvoke(new Action(delegate { window.Close(); }));
                    Assert(owner.Join(5000), "Synthetic window did not shut down");
                }
            }
            Console.WriteLine("PASS: managed UIA, real synthetic controls, password handling, redaction, header recognition, node limits and coverage labels.");
            return 0;
        }
    }
}
