using AlibreAddOn;
using AlibreX;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using IStream = System.Runtime.InteropServices.ComTypes.IStream;
using MessageBox = System.Windows.MessageBox;

namespace AlibreAddOnAssembly
{
    public static class AlibreAddOn
    {
        private static IADRoot AlibreRoot { get; set; }
        private static AddOnRibbon TheAddOnInterface { get; set; }
        private static ScriptRunner PythonRunner { get; set; }

        public static void AddOnLoad(IntPtr hwnd, IAutomationHook pAutomationHook, IntPtr unused)
        {
            AlibreRoot = (IADRoot)pAutomationHook.Root;
            PythonRunner = new ScriptRunner(AlibreRoot);
            TheAddOnInterface = new AddOnRibbon(AlibreRoot);
        }

        public static void AddOnUnload(IntPtr hwnd, bool forceUnload, ref bool cancel, int reserved1, int reserved2)
        {
            TheAddOnInterface = null;
            PythonRunner = null;
            AlibreRoot = null;
        }

        public static void AddOnInvoke(IntPtr pAutomationHook, string sessionName, bool isLicensed, int reserved1, int reserved2) { }
        public static IAlibreAddOn GetAddOnInterface() => TheAddOnInterface;
        public static ScriptRunner GetScriptRunner() => PythonRunner;
    }

    public class AddOnRibbon : IAlibreAddOn
    {
        private readonly MenuManager _menuManager;
        private readonly IADRoot _alibreRoot;

        public AddOnRibbon(IADRoot alibreRoot)
        {
            _alibreRoot = alibreRoot;
            _menuManager = new MenuManager();
        }

        public int RootMenuItem => _menuManager.GetRootMenuItem().Id;
        public bool HasSubMenus(int menuID) => _menuManager.GetMenuItemById(menuID)?.SubItems.Count > 0;
        public Array SubMenuItems(int menuID) => _menuManager.GetMenuItemById(menuID)?.SubItems.Select(subItem => subItem.Id).ToArray();
        public string MenuItemText(int menuID) => _menuManager.GetMenuItemById(menuID)?.Text;
        public string MenuItemToolTip(int menuID) => _menuManager.GetMenuItemById(menuID)?.ToolTip;

        // Icon functionality disabled: always returns null
        public string MenuIcon(int menuID) => null;

        public IAlibreAddOnCommand InvokeCommand(int menuID, string sessionIdentifier)
        {
            var session = _alibreRoot.Sessions.Item(sessionIdentifier);
            var menuItem = _menuManager.GetMenuItemById(menuID);
            return menuItem?.Command?.Invoke(session);
        }

        public ADDONMenuStates MenuItemState(int menuID, string sessionIdentifier) => ADDONMenuStates.ADDON_MENU_ENABLED;
        public bool PopupMenu(int menuID) => false;
        public bool HasPersistentDataToSave(string sessionIdentifier) => false;
        public void SaveData(IStream pCustomData, string sessionIdentifier) { }
        public void LoadData(IStream pCustomData, string sessionIdentifier) { }
        public bool UseDedicatedRibbonTab() => false;
        void IAlibreAddOn.setIsAddOnLicensed(bool isLicensed) { }

        public void LoadData(global::AlibreAddOn.IStream pCustomData, string sessionIdentifier)
        {
            throw new NotImplementedException();
        }

        public void SaveData(global::AlibreAddOn.IStream pCustomData, string sessionIdentifier)
        {
            throw new NotImplementedException();
        }
    }

    public class MenuItem
    {
        public int Id { get; set; }
        public string Text { get; set; }
        public string ToolTip { get; set; }

        public string Icon { get; set; }

        public Func<IADSession, IAlibreAddOnCommand> Command { get; set; }
        public List<MenuItem> SubItems { get; set; } = new List<MenuItem>();

        public MenuItem(int id, string text, string toolTip = "", string icon = null)
        {
            Id = id;
            Text = text;
            ToolTip = toolTip;
            Icon = null;
        }

        public void AddSubItem(MenuItem subItem) => SubItems.Add(subItem);

        public IAlibreAddOnCommand AboutCmd(IADSession session)
        {
            MessageBox.Show("Alibre Neutralizer\r\nParametric Bulk-Exporter for Alibre Design\r\nhttps://github.com/k4kfh/alibre-neutralizer\r\n");
            return null;
        }
    }

    public class MenuManager
    {
        private readonly MenuItem _rootMenuItem;
        private readonly Dictionary<int, MenuItem> _menuItems = new Dictionary<int, MenuItem>();

        public MenuManager()
        {
            _rootMenuItem = new MenuItem(401, "Alibre Neutralizer", "Alibre Neutralizer");
            BuildMenus();
            RegisterMenuItem(_rootMenuItem);
        }

        private void BuildMenus()
        {
            var runItem = new MenuItem(10000, "Run Alibre Neutralizer", "Run the Alibre Neutralizer bulk-export tool");
            runItem.Command = (session) =>
            {
                AlibreAddOn.GetScriptRunner()?.ExecuteScript(session, "alibre-neutralizer.py");
                return null;
            };
            _rootMenuItem.AddSubItem(runItem);

            var aboutItem = new MenuItem(9090, "About", "https://github.com/k4kfh/alibre-neutralizer");
            aboutItem.Command = aboutItem.AboutCmd;
            _rootMenuItem.AddSubItem(aboutItem);
        }

        private void RegisterMenuItem(MenuItem menuItem)
        {
            _menuItems[menuItem.Id] = menuItem;
            foreach (var subItem in menuItem.SubItems)
                RegisterMenuItem(subItem);
        }
        public MenuItem GetMenuItemById(int id) => _menuItems.TryGetValue(id, out var menuItem) ? menuItem : null;
        public MenuItem GetRootMenuItem() => _rootMenuItem;
    }
    public class ScriptRunner
    {
        private ScriptEngine _engine;
        private ScriptScope _scope;
        private readonly IADRoot _alibreRoot;

        public ScriptRunner(IADRoot alibreRoot)
        {
            _alibreRoot = alibreRoot;
            InitializePythonEngine();
        }

        private void InitializePythonEngine()
        {
            try
            {
                var options = new Dictionary<string, object>();
                options["LightweightScopes"] = true;
                var engine = Python.CreateEngine(options);

                string alibreInstallPath = Assembly.GetAssembly(typeof(IADRoot)).Location
                    .Replace("\\Program\\AlibreX.dll", "");

                var searchPaths = engine.GetSearchPaths();
                searchPaths.Add(Path.Combine(alibreInstallPath, "Program"));
                searchPaths.Add(Path.Combine(alibreInstallPath, "Program", "Addons", "AlibreScript", "PythonLib"));
                searchPaths.Add(Path.Combine(alibreInstallPath, "Program", "Addons", "AlibreScript"));
                searchPaths.Add(Path.Combine(alibreInstallPath, "Program", "Addons", "AlibreScript", "PythonLib", "site-packages"));

                string sitePackagesPath = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    "PythonLib", "site-packages");
                if (Directory.Exists(sitePackagesPath))
                    searchPaths.Add(sitePackagesPath);

                engine.SetSearchPaths(searchPaths);

                _engine = engine;

                var scope = engine.CreateScope();
                scope.SetVariable("ScriptFileName", "");
                scope.SetVariable("ScriptFolder", "");
                scope.SetVariable("SessionIdentifier", "");
                scope.SetVariable("WizoScriptVersion", 347013);
                scope.SetVariable("AlibreScriptVersion", 347013);
                scope.SetVariable("Arguments", new List<string>());

                if (_alibreRoot != null)
                    scope.SetVariable("AlibreRoot", _alibreRoot);

                string addOnDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string pythonLibPath = Path.Combine(addOnDirectory, "PythonLib", "site-packages");

                string setupCode =
                    "import sys\n" +
                    "import clr\n" +
                    "import System\n" +
                    "\n" +
                    "# Ensure we're using the correct IronPython runtime\n" +
                    "clr.AddReference('IronPython')\n" +
                    "clr.AddReference('Microsoft.Scripting')\n" +
                    "\n" +
                    "# Load core Alibre assemblies (matching AlibreScript)\n" +
                    "clr.AddReference('AlibreX')\n" +
                    "clr.AddReference('AlibreScriptAddOn')\n" +
                    "\n" +
                    "# Import needed namespaces\n" +
                    "from System.Runtime.InteropServices import Marshal\n" +
                    "import AlibreX\n" +
                    "from AlibreX import *\n" +
                    "\n" +
                    $"if r'{addOnDirectory}' not in sys.path: sys.path.append(r'{addOnDirectory}')\n" +
                    $"if r'{pythonLibPath}' not in sys.path: sys.path.append(r'{pythonLibPath}')\n" +
                    "if 'PythonLib/site-packages' not in sys.path: sys.path.append('PythonLib/site-packages')\n" +
                    "\n" +
                    "try:\n" +
                    "    clr.AddReference('IronPython.SQLite.dll')\n" +
                    "except:\n" +
                    "    pass\n" +
                    "\n" +
                    "from AlibreScript.API import *\n" +
                    "\n" +
                    "# Set default units\n" +
                    "try:\n" +
                    "    Units.Current = UnitTypes.Millimeters\n" +
                    "except:\n" +
                    "    pass\n" +
                    "\n" +
                    "# Helper function to convert Python list to .NET List (recursive)\n" +
                    "def ToList(python_list, item_type=None, recursive=True):\n" +
                    "    from System.Collections.Generic import List\n" +
                    "    if item_type is None:\n" +
                    "        item_type = object\n" +
                    "    net_list = List[item_type]()\n" +
                    "    for item in python_list:\n" +
                    "        if recursive and isinstance(item, list):\n" +
                    "            net_list.Add(ToList(item, object, True))\n" +
                    "        else:\n" +
                    "            net_list.Add(item)\n" +
                    "    return net_list\n" +
                    "\n" +
                    "try:\n" +
                    "    alibre = Marshal.GetActiveObject('AlibreX.AutomationHook')\n" +
                    "    root = alibre.Root\n" +
                    "except:\n" +
                    "    if 'AlibreRoot' in dir() and AlibreRoot is not None:\n" +
                    "        root = AlibreRoot\n" +
                    "    else:\n" +
                    "        root = None\n" +
                    "\n" +
                    "# Define CurrentPart and CurrentAssembly as callables (matching AlibreScript)\n" +
                    "def CurrentPart():\n" +
                    "    if CurrentSession is None:\n" +
                    "        print('WARNING: CurrentSession is None - no active part session')\n" +
                    "        return None\n" +
                    "    try:\n" +
                    "        from AlibreScript.API import Part as PartClass\n" +
                    "        part = PartClass(CurrentSession)\n" +
                    "        return part\n" +
                    "    except Exception as e1:\n" +
                    "        try:\n" +
                    "            part = PartClass(SessionIdentifier, False)\n" +
                    "            return part\n" +
                    "        except Exception as e2:\n" +
                    "            print('ERROR in CurrentPart():')\n" +
                    "            print('  First attempt (CurrentSession):', str(e1))\n" +
                    "            print('  Second attempt (SessionIdentifier):', str(e2))\n" +
                    "    return None\n" +
                    "\n" +
                    "def CurrentAssembly():\n" +
                    "    if CurrentSession is None:\n" +
                    "        print('WARNING: CurrentSession is None - no active assembly session')\n" +
                    "        return None\n" +
                    "    try:\n" +
                    "        from AlibreScript.API import Assembly as AssemblyClass\n" +
                    "        assembly = AssemblyClass(CurrentSession)\n" +
                    "        return assembly\n" +
                    "    except Exception as e1:\n" +
                    "        try:\n" +
                    "            assembly = AssemblyClass(SessionIdentifier, False)\n" +
                    "            return assembly\n" +
                    "        except Exception as e2:\n" +
                    "            print('ERROR in CurrentAssembly():', str(e1), str(e2))\n" +
                    "    return None\n" +
                    "\n" +
                    "def CurrentParts():\n" +
                    "    parts = []\n" +
                    "    part = CurrentPart()\n" +
                    "    if part is not None:\n" +
                    "        parts.append(part)\n" +
                    "    return parts\n" +
                    "\n" +
                    "def CurrentAssemblies():\n" +
                    "    assemblies = []\n" +
                    "    assembly = CurrentAssembly()\n" +
                    "    if assembly is not None:\n" +
                    "        assemblies.append(assembly)\n" +
                    "    return assemblies\n" +
                    "\n" +
                    "# Global Windows instance (will be set by .NET with proper parent form)\n" +
                    "_PreCreatedWindowsInstance = None\n" +
                    "\n" +
                    "# Helper function Windows() matching AlibreScript\n" +
                    "def Windows():\n" +
                    "    global _PreCreatedWindowsInstance\n" +
                    "    if _PreCreatedWindowsInstance is not None:\n" +
                    "        return _PreCreatedWindowsInstance\n" +
                    "    from AlibreScript.API import Windows as Win\n" +
                    "    return Win(SessionIdentifier, '', None)\n" +
                    "\n" +
                    "# Tracing functions\n" +
                    "def StartTracing(path):\n" +
                    "    try:\n" +
                    "        from AlibreScript.API import Trace\n" +
                    "        Trace.Start(path)\n" +
                    "    except:\n" +
                    "        pass\n" +
                    "\n" +
                    "def StopTracing():\n" +
                    "    try:\n" +
                    "        from AlibreScript.API import Trace\n" +
                    "        Trace.Stop()\n" +
                    "    except:\n" +
                    "        pass\n" +
                    "\n" +
                    "# CSharp function\n" +
                    "def CSharp():\n" +
                    "    from AlibreScript.API import CSharp as CS\n" +
                    "    return CS(SessionIdentifier, '', None)\n" +
                    "\n" +
                    "# Convenience helper functions\n" +
                    "def InfoDialog(message, title='Info'):\n" +
                    "    Windows().InfoDialog(message, title)\n" +
                    "\n" +
                    "def ErrorDialog(message, title='Error'):\n" +
                    "    Windows().ErrorDialog(message, title)\n" +
                    "\n" +
                    "# Helper to wrap Python callbacks as .NET Action delegates\n" +
                    "def WrapCallback(func):\n" +
                    "    if func is None:\n" +
                    "        return None\n" +
                    "    from System import Action\n" +
                    "    return func\n";

                engine.Execute(setupCode, scope);

                _scope = scope;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing IronPython engine: {ex.Message}", "Error");
            }
        }

        public void ExecuteScript(IADSession session, string mainScriptFileName)
        {
            try
            {
                string addOnDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string scriptsPath = Path.Combine(addOnDirectory, "Scripts");
                string mainScriptPath = Path.Combine(scriptsPath, mainScriptFileName);
                if (!File.Exists(mainScriptPath))
                {
                    MessageBox.Show($"Error: Script not found.\nPath: {mainScriptPath}", "Script Error");
                    return;
                }

                // Update session-specific variables before each script execution
                string sessionId;
                try
                {
                    sessionId = session?.Identifier ?? Guid.NewGuid().ToString();
                }
                catch
                {
                    sessionId = Guid.NewGuid().ToString();
                }

                _scope.SetVariable("ScriptFileName", mainScriptFileName);
                _scope.SetVariable("ScriptFolder", scriptsPath);
                _scope.SetVariable("SessionIdentifier", sessionId);
                _scope.SetVariable("CurrentSession", session);

                if (session != null)
                    _scope.SetVariable("Session", session);

                // Try to create a Windows instance with proper parent form
                try
                {
                    var asm = System.Reflection.Assembly.Load("AlibreScriptAddOn");
                    var windowsType = asm.GetType("AlibreScript.API.Windows");
                    var windowsInstance = Activator.CreateInstance(windowsType, sessionId, "", (object)null);
                    if (windowsInstance != null)
                        _scope.SetVariable("_PreCreatedWindowsInstance", windowsInstance);
                }
                catch { }

                _engine.ExecuteFile(mainScriptPath, _scope);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while running the script:\n{ex}", "Python Execution Error");
            }
        }
    }
}