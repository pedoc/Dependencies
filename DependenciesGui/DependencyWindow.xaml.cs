using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.ComponentModel;
using System.Windows.Input;
using System.Diagnostics;
using System.Linq;
using System.Windows.Data;
using System.Windows.Forms;
using Microsoft.Win32;
using Mono.Cecil;
using Dependencies.ClrPh;
using FastDeepCloner;
using Application = System.Windows.Application;
using Clipboard = System.Windows.Clipboard;
using MessageBox = System.Windows.MessageBox;
using TreeView = System.Windows.Controls.TreeView;

namespace Dependencies
{
    /// <summary>
    /// ImportContext : Describe an import module parsed from a PE.
    /// Only used during the dependency tree building phase
    /// </summary>
    public struct ImportContext
    {
        // Import "identifier" 
        public string ModuleName;

        // Return how the module was found (NOT_FOUND otherwise)
        public ModuleSearchStrategy ModuleLocation;

        // If found, set the filepath and parsed PE, otherwise it's null
        public string PeFilePath;
        public PE PeProperties;

        // Some imports are from api sets
        public bool IsApiSet;
        public string ApiSetModuleName;

        // module flag attributes
        public ModuleFlag Flags;
    }


    /// <summary>
    /// Dependency tree building behaviour.
    /// A full recursive dependency tree can be memory intensive, therefore the
    /// choice is left to the user to override the default behaviour.
    /// </summary>
    public class TreeBuildingBehaviour : IValueConverter
    {
        public enum DependencyTreeBehaviour
        {
            ChildOnly,
            RecursiveOnlyOnDirectImports,
            Recursive,
        }

        public static DependencyTreeBehaviour GetGlobalBehaviour()
        {
            return (DependencyTreeBehaviour)(new TreeBuildingBehaviour()).Convert(
                Properties.Settings.Default.TreeBuildBehaviour,
                null, // targetType
                null, // parameter
                null // System.Globalization.CultureInfo
            );
        }

        #region TreeBuildingBehaviour.IValueConverter_contract

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string strBehaviour = (string)value;

            switch (strBehaviour)
            {
                default:
                case "ChildOnly":
                    return DependencyTreeBehaviour.ChildOnly;
                case "RecursiveOnlyOnDirectImports":
                    return DependencyTreeBehaviour.RecursiveOnlyOnDirectImports;
                case "Recursive":
                    return DependencyTreeBehaviour.Recursive;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            DependencyTreeBehaviour behaviour = (DependencyTreeBehaviour)value;

            switch (behaviour)
            {
                default:
                case DependencyTreeBehaviour.ChildOnly:
                    return "ChildOnly";
                case DependencyTreeBehaviour.RecursiveOnlyOnDirectImports:
                    return "RecursiveOnlyOnDirectImports";
                case DependencyTreeBehaviour.Recursive:
                    return "Recursive";
            }
        }

        #endregion TreeBuildingBehaviour.IValueConverter_contract
    }

    /// <summary>
    /// Dependency tree building behaviour.
    /// A full recursive dependency tree can be memory intensive, therefore the
    /// choice is left to the user to override the default behaviour.
    /// </summary>
    public class BinaryCacheOption : IValueConverter
    {
        [TypeConverter(typeof(EnumToStringUsingDescription))]
        public enum BinaryCacheOptionValue
        {
            [Description("No (faster, but locks dll until Dependencies is closed)")]
            No = 0,

            [Description("Yes (prevents file locking issues)")]
            Yes = 1
        }

        public static BinaryCacheOptionValue GetGlobalBehaviour()
        {
            return (BinaryCacheOptionValue)(new BinaryCacheOption()).Convert(
                Properties.Settings.Default.BinaryCacheOptionValue,
                null, // targetType
                null, // parameter
                null // System.Globalization.CultureInfo
            );
        }

        #region BinaryCacheOption.IValueConverter_contract

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            bool strOption = (bool)value;

            switch (strOption)
            {
                default:
                case true:
                    return BinaryCacheOptionValue.Yes;
                case false:
                    return BinaryCacheOptionValue.No;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            BinaryCacheOptionValue behaviour = (BinaryCacheOptionValue)(int)value;

            switch (behaviour)
            {
                default:
                case BinaryCacheOptionValue.Yes:
                    return true;
                case BinaryCacheOptionValue.No:
                    return false;
            }
        }

        #endregion BinaryCacheOption.IValueConverter_contract
    }

    public class EnumToStringUsingDescription : TypeConverter
    {
        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
        {
            return (sourceType.Equals(typeof(Enum)));
        }

        public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType)
        {
            return (destinationType.Equals(typeof(String)));
        }

        public override object ConvertFrom(ITypeDescriptorContext context, System.Globalization.CultureInfo culture,
            object value)
        {
            return base.ConvertFrom(context, culture, value);
        }

        public override object ConvertTo(ITypeDescriptorContext context, System.Globalization.CultureInfo culture,
            object value, Type destinationType)
        {
            if (!destinationType.Equals(typeof(String)))
            {
                throw new ArgumentException("Can only convert to string.", "destinationType");
            }

            if (!value.GetType().BaseType.Equals(typeof(Enum)))
            {
                throw new ArgumentException("Can only convert an instance of enum.", "value");
            }

            string name = value.ToString();
            object[] attrs =
                value.GetType().GetField(name).GetCustomAttributes(typeof(DescriptionAttribute), false);
            return (attrs.Length > 0) ? ((DescriptionAttribute)attrs[0]).Description : name;
        }
    }

    /// <summary>
    /// User context of every dependency tree node.
    /// </summary>
    public struct DependencyNodeContext
    {
        public DependencyNodeContext(DependencyNodeContext other)
        {
            ModuleInfo = other.ModuleInfo;
            IsDummy = other.IsDummy;
        }

        /// <summary>
        /// We use a WeakReference to point towars a DisplayInfoModule
        /// in order to reduce memory allocations.
        /// </summary>
        public WeakReference ModuleInfo;

        /// <summary>
        /// Depending on the dependency tree behaviour, we may have to
        /// set up "dummy" nodes in order for the parent to display the ">" button.
        /// Those dummy are usually destroyed when their parents is expandend and imports resolved.
        /// </summary>
        public bool IsDummy;
    }

    /// <summary>
    /// Deprendency Tree custom node. It's DataContext is a DependencyNodeContext struct
    /// </summary>
    public class ModuleTreeViewItem : TreeViewItem, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public ModuleTreeViewItem()
        {
            _importsVerified = false;
            ParentModule = null;
            Properties.Settings.Default.PropertyChanged += ModuleTreeViewItem_PropertyChanged;
        }

        public ModuleTreeViewItem(ModuleTreeViewItem parent)
        {
            _importsVerified = false;
            ParentModule = parent;
            Properties.Settings.Default.PropertyChanged += ModuleTreeViewItem_PropertyChanged;
        }

        public ModuleTreeViewItem(ModuleTreeViewItem other, ModuleTreeViewItem parent)
        {
            _importsVerified = false;
            ParentModule = parent;
            DataContext = new DependencyNodeContext((DependencyNodeContext)other.DataContext);
            Properties.Settings.Default.PropertyChanged += ModuleTreeViewItem_PropertyChanged;
        }

        #region PropertyEventHandlers

        public virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void ModuleTreeViewItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "FullPath")
            {
                Header = (object)GetTreeNodeHeaderName(Properties.Settings.Default.FullPath);
            }
        }

        #endregion PropertyEventHandlers

        #region Getters

        public string GetTreeNodeHeaderName(bool fullPath)
        {
            return ((((DependencyNodeContext)DataContext).ModuleInfo.Target as DisplayModuleInfo)!).ModuleName;
        }

        public string ModuleFilePath =>
            ((((DependencyNodeContext)DataContext).ModuleInfo.Target as DisplayModuleInfo)!).Filepath;

        public ModuleTreeViewItem ParentModule { get; }


        public ModuleFlag Flags => ModuleInfo.Flags;

        private bool _hasError;

        public bool HasErrors
        {
            get
            {
                if (!_importsVerified)
                {
                    _hasError = VerifyModuleImports();
                    _importsVerified = true;

                    // Update tooltip only once some basic checks are done
                    ToolTip = ModuleInfo.Status;
                }

                // propagate error for parent
                if (_hasError)
                {
                    ModuleTreeViewItem parentModule = this.ParentModule;
                    if (parentModule != null)
                    {
                        parentModule.HasChildErrors = true;
                    }
                }

                return _hasError;
            }

            set
            {
                if (value == _hasError) return;
                _hasError = value;
                OnPropertyChanged("HasErrors");
            }
        }


        public string Tooltip => ModuleInfo.Status;

        public bool HasChildErrors
        {
            get => _hasChildErrors;
            set
            {
                if (value)
                {
                    ModuleInfo.Flags |= ModuleFlag.ChildrenError;
                }
                else
                {
                    ModuleInfo.Flags &= ~ModuleFlag.ChildrenError;
                }

                ToolTip = ModuleInfo.Status;
                _hasChildErrors = true;
                OnPropertyChanged("HasChildErrors");

                // propagate error for parent
                ModuleTreeViewItem parentModule = this.ParentModule;
                if (parentModule != null)
                {
                    parentModule.HasChildErrors = true;
                }
            }
        }

        public DisplayModuleInfo ModuleInfo =>
            (((DependencyNodeContext)DataContext).ModuleInfo.Target as DisplayModuleInfo);


        private bool VerifyModuleImports()
        {
            // current module has issues
            if ((Flags & (ModuleFlag.NotFound | ModuleFlag.MissingImports | ModuleFlag.ChildrenError)) != 0)
            {
                return true;
            }

            // no parent : it's probably the root item
            ModuleTreeViewItem parentModule = this.ParentModule;
            if (parentModule == null)
            {
                return false;
            }

            // Check we have any imports issues
            foreach (PeImportDll dllImport in parentModule.ModuleInfo.Imports)
            {
                if (dllImport.Name != ModuleInfo._Name)
                    continue;


                List<Tuple<PeImport, bool>> resolvedImports = BinaryCache.LookupImports(dllImport, ModuleInfo.Filepath);
                if (resolvedImports.Count == 0)
                {
                    return true;
                }

                foreach (var import in resolvedImports)
                {
                    if (!import.Item2)
                    {
                        return true;
                    }
                }
            }


            return false;
        }

        #endregion Getters


        #region Commands

        public RelayCommand OpenPeviewerCommand
        {
            get
            {
                if (_openPeviewerCommand == null)
                {
                    _openPeviewerCommand = new RelayCommand((param) => OpenPeviewer((object)param));
                }

                return _openPeviewerCommand;
            }
        }

        public bool OpenPeviewer(object context)
        {
            string programPath = Properties.Settings.Default.PeViewerPath;
            Process peviewerProcess = new Process();

            if (context == null)
            {
                return false;
            }

            if (!File.Exists(programPath))
            {
                MessageBox.Show(String.Format("{0:s} file could not be found !", programPath));
                return false;
            }

            string filepath = ModuleFilePath;
            if (filepath == null)
            {
                return false;
            }

            peviewerProcess.StartInfo.FileName = String.Format("\"{0:s}\"", programPath);
            peviewerProcess.StartInfo.Arguments = String.Format("\"{0:s}\"", filepath);
            return peviewerProcess.Start();
        }

        public RelayCommand OpenNewAppCommand
        {
            get
            {
                if (_openNewAppCommand == null)
                {
                    _openNewAppCommand = new RelayCommand((param) =>
                    {
                        string filepath = ModuleFilePath;
                        if (filepath == null)
                        {
                            return;
                        }

                        Process otherDependenciesProcess = new Process();
                        otherDependenciesProcess.StartInfo.FileName = System.Windows.Forms.Application.ExecutablePath;
                        otherDependenciesProcess.StartInfo.Arguments = String.Format("\"{0:s}\"", filepath);
                        otherDependenciesProcess.Start();
                    });
                }

                return _openNewAppCommand;
            }
        }

        #endregion // Commands 

        private RelayCommand _openPeviewerCommand;
        private RelayCommand _openNewAppCommand;
        private bool _importsVerified;
        private bool _hasChildErrors;
    }


    /// <summary>
    /// Dependemcy tree analysis window for a given PE.
    /// </summary>
    public partial class DependencyWindow : TabItem
    {
        PE _pe;
        public string RootFolder;
        public string WorkingDirectory;
        string _filename;
        PhSymbolProvider _symPrv;
        SxsEntries _sxsEntriesCache;
        ApiSetSchema _apiSetmapCache;
        ModulesCache _processedModulesCache;
        DisplayModuleInfo _selectedModule;
        bool _displayWarning;

        public List<string> CustomSearchFolders;

        #region PublicAPI

        public DependencyWindow(String filename, List<string> customSearchFolders = null)
        {
            InitializeComponent();

            if (customSearchFolders != null)
            {
                this.CustomSearchFolders = customSearchFolders;
            }
            else
            {
                this.CustomSearchFolders = new List<string>();
            }

            this._filename = filename;
            WorkingDirectory = Path.GetDirectoryName(this._filename);
            InitializeView();
        }

        private static PE LoadPe(string filename)
        {
            return (Application.Current as App)!.LoadBinary(filename);
        }

        public void InitializeView()
        {
            if (!NativeFile.Exists(_filename))
            {
                MessageBox.Show(
                    String.Format("{0:s} is not present on the disk", _filename),
                    "Invalid PE",
                    MessageBoxButton.OK
                );

                return;
            }

            _pe = LoadPe(_filename);
            if (_pe == null || !_pe.LoadSuccessful)
            {
                MessageBox.Show(
                    String.Format("{0:s} is not a valid PE-COFF file", _filename),
                    "Invalid PE",
                    MessageBoxButton.OK
                );

                return;
            }

            _symPrv = new PhSymbolProvider();
            RootFolder = Path.GetDirectoryName(_filename);
            _sxsEntriesCache = SxsManifest.GetSxsEntries(_pe);
            _processedModulesCache = new ModulesCache();
            _apiSetmapCache = Phlib.GetApiSetSchema();
            _selectedModule = null;
            _displayWarning = false;

            // TODO : Find a way to properly bind commands instead of using this hack
            ModulesList.Items.Clear();
            ModulesList.DoFindModuleInTreeCommand = DoFindModuleInTree;
            ModulesList.ConfigureSearchOrderCommand = ConfigureSearchOrderCommand;

            var rootFilename = Path.GetFileName(_filename);
            var rootModule = new DisplayModuleInfo(rootFilename, _pe, ModuleSearchStrategy.ROOT);
            _processedModulesCache.Add(new ModuleCacheKey(rootFilename, _filename), rootModule);
            ModulesList.AddModule(rootModule);

            ModuleTreeViewItem treeNode = new ModuleTreeViewItem();
            DependencyNodeContext childTreeInfoContext = new DependencyNodeContext()
            {
                ModuleInfo = new WeakReference(rootModule),
                IsDummy = false
            };

            treeNode.DataContext = childTreeInfoContext;
            treeNode.Header = treeNode.GetTreeNodeHeaderName(Properties.Settings.Default.FullPath);
            treeNode.IsExpanded = true;

            DllTreeView.Items.Clear();
            DllTreeView.Items.Add(treeNode);

            TreeBuildingBehaviour.DependencyTreeBehaviour settingTreeBehaviour =
                TreeBuildingBehaviour.GetGlobalBehaviour();
            // Recursively construct tree of dll imports
            ConstructDependencyTree(treeNode, _pe, settingTreeBehaviour);
        }

        #endregion PublicAPI


        #region TreeConstruction

        private ImportContext ResolveImport(PeImportDll dllImport)
        {
            return ResolveImport(dllImport, _pe, _sxsEntriesCache, CustomSearchFolders, WorkingDirectory);
        }

        private static ImportContext ResolveImport(
            PeImportDll dllImport,
            PE rootPe,
            SxsEntries sxsCache,
            List<string> customSearchFolders,
            string workingDirectory)
        {
            ImportContext importModule = new ImportContext();

            importModule.PeFilePath = null;
            importModule.PeProperties = null;
            importModule.ModuleName = dllImport.Name;
            importModule.ApiSetModuleName = null;
            importModule.Flags = 0;
            if (dllImport.IsDelayLoad())
            {
                importModule.Flags |= ModuleFlag.DelayLoad;
            }

            Tuple<ModuleSearchStrategy, PE> resolvedModule = BinaryCache.ResolveModule(
                rootPe,
                dllImport.Name,
                sxsCache,
                customSearchFolders,
                workingDirectory
            );

            importModule.ModuleLocation = resolvedModule.Item1;
            if (importModule.ModuleLocation != ModuleSearchStrategy.NOT_FOUND)
            {
                importModule.PeProperties = resolvedModule.Item2;

                if (resolvedModule.Item2 != null)
                {
                    importModule.PeFilePath = resolvedModule.Item2.Filepath;
                    foreach (var import in BinaryCache.LookupImports(dllImport, importModule.PeFilePath))
                    {
                        if (!import.Item2)
                        {
                            importModule.Flags |= ModuleFlag.MissingImports;
                            break;
                        }
                    }
                }
            }
            else
            {
                importModule.Flags |= ModuleFlag.NotFound;
            }

            // special case for apiset schema
            importModule.IsApiSet = (importModule.ModuleLocation == ModuleSearchStrategy.ApiSetSchema);
            if (importModule.IsApiSet)
            {
                importModule.Flags |= ModuleFlag.ApiSet;
                importModule.ApiSetModuleName = BinaryCache.LookupApiSetLibrary(dllImport.Name);

                if (dllImport.Name.StartsWith("ext-"))
                {
                    importModule.Flags |= ModuleFlag.ApiSetExt;
                }
            }

            return importModule;
        }

        private static void TriggerWarningOnAppvIsvImports(string dllImportName, ref bool displayWarning)
        {
            if (String.Compare(dllImportName, "AppvIsvSubsystems32.dll", StringComparison.OrdinalIgnoreCase) == 0 ||
                String.Compare(dllImportName, "AppvIsvSubsystems64.dll", StringComparison.OrdinalIgnoreCase) == 0)
            {
                if (!displayWarning)
                {
                    MessageBoxResult result = MessageBox.Show(
                        "This binary use the App-V containerization technology which fiddle with search directories and PATH env in ways Dependencies can't handle.\n\nFollowing results are probably not quite exact.",
                        "App-V ISV disclaimer"
                    );

                    displayWarning = true; // prevent the same warning window to popup several times
                }
            }
        }

        private void TriggerWarningOnAppvIsvImports(string dllImportName)
        {
            TriggerWarningOnAppvIsvImports(dllImportName, ref _displayWarning);
        }

        private void ProcessAppInitDlls(
            Dictionary<string, ImportContext> newTreeContexts,
            PE analyzedPe,
            ImportContext importModule)
        {
            ProcessAppInitDlls(_pe, newTreeContexts, analyzedPe, importModule, _sxsEntriesCache, CustomSearchFolders,
                WorkingDirectory);
        }

        private static void ProcessAppInitDlls(
            PE rootPe,
            Dictionary<string, ImportContext> newTreeContexts,
            PE analyzedPe,
            ImportContext importModule,
            SxsEntries sxsEntriesCache,
            List<string> customSearchFolders,
            string workingDirectory
        )
        {
            List<PeImportDll> peImports = analyzedPe.GetImports();

            // only user32 triggers appinit dlls
            string user32Filepath = Path.Combine(FindPe.GetSystemPath(rootPe), "user32.dll");
            if (importModule.PeFilePath != user32Filepath)
            {
                return;
            }

            string appInitRegistryKey =
                (rootPe.IsArm32Dll()) ? "SOFTWARE\\WowAA32Node\\Microsoft\\Windows NT\\CurrentVersion\\Windows" :
                (rootPe.IsWow64Dll()) ? "SOFTWARE\\Wow6432Node\\Microsoft\\Windows NT\\CurrentVersion\\Windows" :
                "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Windows";

            // Opening registry values
            RegistryKey localKey =
                RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            localKey = localKey.OpenSubKey(appInitRegistryKey);
            int loadAppInitDlls = (int)localKey.GetValue("LoadAppInit_DLLs", 0);
            string appInitDlls = (string)localKey.GetValue("AppInit_DLLs", "");
            if (loadAppInitDlls == 0 || String.IsNullOrEmpty(appInitDlls))
            {
                return;
            }

            // Extremely crude parser. TODO : Add support for quotes wrapped paths with spaces
            foreach (var appInitDll in appInitDlls.Split(' '))
            {
                Debug.WriteLine("AppInit loading " + appInitDll);

                // Do not process twice the same imported module
                if (null != peImports.Find(module => module.Name == appInitDll))
                {
                    continue;
                }

                if (newTreeContexts.ContainsKey(appInitDll))
                {
                    continue;
                }

                ImportContext appInitImportModule = new ImportContext();
                appInitImportModule.PeFilePath = null;
                appInitImportModule.PeProperties = null;
                appInitImportModule.ModuleName = appInitDll;
                appInitImportModule.ApiSetModuleName = null;
                appInitImportModule.Flags = 0;
                appInitImportModule.ModuleLocation = ModuleSearchStrategy.AppInitDLL;


                Tuple<ModuleSearchStrategy, PE> resolvedAppInitModule = BinaryCache.ResolveModule(
                    rootPe,
                    appInitDll,
                    sxsEntriesCache,
                    customSearchFolders,
                    workingDirectory
                );
                if (resolvedAppInitModule.Item1 != ModuleSearchStrategy.NOT_FOUND)
                {
                    appInitImportModule.PeProperties = resolvedAppInitModule.Item2;
                    appInitImportModule.PeFilePath = resolvedAppInitModule.Item2.Filepath;
                }
                else
                {
                    appInitImportModule.Flags |= ModuleFlag.NotFound;
                }

                newTreeContexts.Add(appInitDll, appInitImportModule);
            }
        }

        private void ProcessClrImports(Dictionary<string, ImportContext> newTreeContexts, PE analyzedPe)
        {
            ProcessClrImports(RootFolder, _pe, newTreeContexts, analyzedPe, _sxsEntriesCache, CustomSearchFolders,
                WorkingDirectory);
        }

        private static void ProcessClrImports(
            string rootFolder,
            PE rootPe,
            Dictionary<string, ImportContext> newTreeContexts,
            PE analyzedPe,
            SxsEntries sxsEntriesCache,
            List<string> customSearchFolders,
            string workingDirectory)
        {
            List<PeImportDll> peImports = analyzedPe.GetImports();

            var resolver = new DefaultAssemblyResolver();
            resolver.AddSearchDirectory(rootFolder);

            // Parse it via cecil
            AssemblyDefinition peAssembly = null;
            try
            {
                peAssembly = AssemblyDefinition.ReadAssembly(analyzedPe.Filepath);
            }
            catch (BadImageFormatException)
            {
                MessageBoxResult result = MessageBox.Show(
                    String.Format(
                        "Cecil could not correctly parse {0:s}, which can happens on .NET Core executables. CLR imports will be not shown",
                        analyzedPe.Filepath),
                    "CLR parsing fail"
                );

                return;
            }

            foreach (var module in peAssembly.Modules)
            {
                // Process CLR referenced assemblies
                foreach (var assembly in module.AssemblyReferences)
                {
                    AssemblyDefinition definition;
                    try
                    {
                        definition = resolver.Resolve(assembly);
                    }
                    catch (AssemblyResolutionException)
                    {
                        ImportContext appInitImportModule = new ImportContext();
                        appInitImportModule.PeFilePath = null;
                        appInitImportModule.PeProperties = null;
                        appInitImportModule.ModuleName = Path.GetFileName(assembly.Name);
                        appInitImportModule.ApiSetModuleName = null;
                        appInitImportModule.Flags = ModuleFlag.ClrReference;
                        appInitImportModule.ModuleLocation = ModuleSearchStrategy.ClrAssembly;
                        appInitImportModule.Flags |= ModuleFlag.NotFound;

                        if (!newTreeContexts.ContainsKey(appInitImportModule.ModuleName))
                        {
                            newTreeContexts.Add(appInitImportModule.ModuleName, appInitImportModule);
                        }

                        continue;
                    }

                    foreach (var assemblyModule in definition.Modules)
                    {
                        Debug.WriteLine("Referenced Assembling loading " + assemblyModule.Name + " : " +
                                        assemblyModule.FileName);

                        // Do not process twice the same imported module
                        if (null != peImports.Find(mod => mod.Name == Path.GetFileName(assemblyModule.FileName)))
                        {
                            continue;
                        }

                        ImportContext appInitImportModule = new ImportContext();
                        appInitImportModule.PeFilePath = null;
                        appInitImportModule.PeProperties = null;
                        appInitImportModule.ModuleName = Path.GetFileName(assemblyModule.FileName);
                        appInitImportModule.ApiSetModuleName = null;
                        appInitImportModule.Flags = ModuleFlag.ClrReference;
                        appInitImportModule.ModuleLocation = ModuleSearchStrategy.ClrAssembly;

                        Tuple<ModuleSearchStrategy, PE> resolvedAppInitModule = BinaryCache.ResolveModule(
                            rootPe,
                            assemblyModule.FileName,
                            workingDirectory
                        );
                        if (resolvedAppInitModule.Item1 != ModuleSearchStrategy.NOT_FOUND)
                        {
                            appInitImportModule.PeProperties = resolvedAppInitModule.Item2;
                            appInitImportModule.PeFilePath = resolvedAppInitModule.Item2.Filepath;
                        }
                        else
                        {
                            appInitImportModule.Flags |= ModuleFlag.NotFound;
                        }

                        if (!newTreeContexts.ContainsKey(appInitImportModule.ModuleName))
                        {
                            newTreeContexts.Add(appInitImportModule.ModuleName, appInitImportModule);
                        }
                    }
                }

                // Process unmanaged dlls for native calls
                foreach (var unmanagedModule in module.ModuleReferences)
                {
                    // some clr dll have a reference to an "empty" dll
                    if (unmanagedModule.Name.Length == 0)
                    {
                        continue;
                    }

                    Debug.WriteLine("Referenced module loading " + unmanagedModule.Name);

                    // Do not process twice the same imported module
                    if (null != peImports.Find(m => m.Name == unmanagedModule.Name))
                    {
                        continue;
                    }


                    ImportContext appInitImportModule = new ImportContext();
                    appInitImportModule.PeFilePath = null;
                    appInitImportModule.PeProperties = null;
                    appInitImportModule.ModuleName = unmanagedModule.Name;
                    appInitImportModule.ApiSetModuleName = null;
                    appInitImportModule.Flags = ModuleFlag.ClrReference;
                    appInitImportModule.ModuleLocation = ModuleSearchStrategy.ClrAssembly;

                    Tuple<ModuleSearchStrategy, PE> resolvedAppInitModule = BinaryCache.ResolveModule(
                        rootPe,
                        unmanagedModule.Name,
                        sxsEntriesCache,
                        customSearchFolders,
                        workingDirectory
                    );
                    if (resolvedAppInitModule.Item1 != ModuleSearchStrategy.NOT_FOUND)
                    {
                        appInitImportModule.PeProperties = resolvedAppInitModule.Item2;
                        appInitImportModule.PeFilePath = resolvedAppInitModule.Item2.Filepath;
                    }

                    if (!newTreeContexts.ContainsKey(appInitImportModule.ModuleName))
                    {
                        newTreeContexts.Add(appInitImportModule.ModuleName, appInitImportModule);
                    }
                }
            }
        }

        private void ProcessPe(
            Dictionary<string, ImportContext> newTreeContexts,
            PE newPe)
        {
            ProcessPe(RootFolder, _pe, newTreeContexts, newPe, _sxsEntriesCache, CustomSearchFolders, WorkingDirectory);
        }

        private static void ProcessPe(
            string rootFolder,
            PE rootPe,
            Dictionary<string, ImportContext> newTreeContexts,
            PE newPe,
            SxsEntries sxsCache,
            List<string> customSearchFolders,
            string workingDirectory)
        {
            List<PeImportDll> peImports = newPe.GetImports();

            foreach (PeImportDll dllImport in peImports)
            {
                // Ignore already processed imports
                if (newTreeContexts.ContainsKey(dllImport.Name))
                {
                    continue;
                }

                // Find Dll in "paths"
                ImportContext importModule =
                    ResolveImport(dllImport, rootPe, sxsCache, customSearchFolders, workingDirectory);
                bool _ = false;
                // add warning for appv isv applications 
                TriggerWarningOnAppvIsvImports(dllImport.Name, ref _);

                newTreeContexts.Add(dllImport.Name, importModule);

                // AppInitDlls are triggered by user32.dll, so if the binary does not import user32.dll they are not loaded.
                ProcessAppInitDlls(rootPe, newTreeContexts, newPe, importModule, sxsCache, customSearchFolders,
                    workingDirectory);
            }

            // This should happen only if this is validated to be a C# assembly
            if (newPe.IsClrDll())
            {
                // We use Mono.Cecil to enumerate its references
                ProcessClrImports(rootFolder, rootPe, newTreeContexts, newPe, sxsCache, customSearchFolders,
                    workingDirectory);
            }
        }

        private class BacklogImport : Tuple<ModuleTreeViewItem, string>
        {
            public BacklogImport(ModuleTreeViewItem node, string filepath)
                : base(node, filepath)
            {
            }
        }

        private void ConstructDependencyTree(
            ModuleTreeViewItem rootNode,
            string filePath,
            TreeBuildingBehaviour.DependencyTreeBehaviour settingTreeBehaviour,
            int recursionLevel = 0,
            Action onRunWorkerCompleted = null)
        {
            PE currentPe = LoadPe(filePath);

            if (null == currentPe)
            {
                return;
            }

            ConstructDependencyTree(rootNode, currentPe, settingTreeBehaviour, recursionLevel, onRunWorkerCompleted);
        }

        private static void ConstructDependencyTreeV2(
            string rootFolder,
            PE rootPe,
            PE newPe,
            Dictionary<string, ImportContext> newTreeContexts,
            SxsEntries sxsCache,
            List<string> customSearchFolders,
            string workingDirectory,
            Dictionary<string, DisplayModuleInfo> modulesCache,
            List<DisplayModuleInfo> moduleList,
            TreeBuildingBehaviour.DependencyTreeBehaviour settingTreeBehaviour,
            int recursionLevel = 0)
        {
            ProcessPe(rootFolder, rootPe, newTreeContexts, newPe, sxsCache, customSearchFolders, workingDirectory);

            List<(ModuleTreeViewItem node, string filepath)> peProcessingBacklog =
                new List<(ModuleTreeViewItem node, string filepath)>();
            foreach (ImportContext newTreeContext in newTreeContexts.Values)
            {
                if (newTreeContext.ModuleLocation == ModuleSearchStrategy.NOT_FOUND) continue;

                string moduleName = newTreeContext.ModuleName;
                var moduleKey = newTreeContext.PeFilePath;

                // Newly seen modules
                if (!modulesCache.ContainsKey(moduleKey))
                {
                    // Missing module "found"
                    if ((newTreeContext.PeFilePath == null) || !NativeFile.Exists(newTreeContext.PeFilePath))
                    {
                        if (newTreeContext.IsApiSet)
                        {
                            modulesCache[moduleKey] =
                                new ApiSetNotFoundModuleInfo(moduleName, newTreeContext.ApiSetModuleName);
                        }
                        else
                        {
                            modulesCache[moduleKey] = new NotFoundModuleInfo(moduleName);
                        }
                    }
                    else
                    {
                        if (newTreeContext.IsApiSet)
                        {
                            var apiSetContractModule = new DisplayModuleInfo(
                                newTreeContext.ApiSetModuleName,
                                newTreeContext.PeProperties,
                                newTreeContext.ModuleLocation,
                                newTreeContext.Flags);
                            var newModule =
                                new ApiSetModuleInfo(newTreeContext.ModuleName, ref apiSetContractModule);

                            modulesCache[moduleKey] = newModule;

                            if (settingTreeBehaviour == TreeBuildingBehaviour.DependencyTreeBehaviour.Recursive)
                            {
                                peProcessingBacklog.Add(
                                    new(
                                        null,
                                        apiSetContractModule.ModuleName));
                            }
                        }
                        else
                        {
                            var newModule = new DisplayModuleInfo(
                                newTreeContext.ModuleName,
                                newTreeContext.PeProperties,
                                newTreeContext.ModuleLocation,
                                newTreeContext.Flags);
                            modulesCache[moduleKey] = newModule;

                            switch (settingTreeBehaviour)
                            {
                                case TreeBuildingBehaviour.DependencyTreeBehaviour.RecursiveOnlyOnDirectImports:
                                    if ((newTreeContext.Flags & ModuleFlag.DelayLoad) == 0)
                                    {
                                        peProcessingBacklog.Add(new(null,
                                            newModule.ModuleName));
                                    }

                                    break;

                                case TreeBuildingBehaviour.DependencyTreeBehaviour.Recursive:
                                    peProcessingBacklog.Add(new(null, newModule.ModuleName));
                                    break;
                            }
                        }
                    }

                    moduleList.Add(modulesCache[moduleKey]);
                }
                else
                {
                }
            }

            bool doProcessNextLevel =
                (settingTreeBehaviour != TreeBuildingBehaviour.DependencyTreeBehaviour.ChildOnly) &&
                (recursionLevel < Properties.Settings.Default.TreeDepth) &&
                peProcessingBacklog.Count > 0;

            if (doProcessNextLevel)
            {
                foreach (var importNode in peProcessingBacklog)
                {
                    var tNewPe = LoadPe(importNode.filepath);
                    ConstructDependencyTreeV2(
                        rootFolder,
                        rootPe,
                        tNewPe,
                        newTreeContexts,
                        sxsCache,
                        customSearchFolders,
                        workingDirectory,
                        modulesCache,
                        moduleList,
                        settingTreeBehaviour,
                        recursionLevel + 1);
                }
            }
        }

        private void ConstructDependencyTree(
            ModuleTreeViewItem rootNode,
            PE currentPe,
            TreeBuildingBehaviour.DependencyTreeBehaviour settingTreeBehaviour,
            int recursionLevel = 0,
            Action onRunWorkerCompleted = null)
        {
            // "Closured" variables (it 's a scope hack really).
            Dictionary<string, ImportContext> newTreeContexts = new Dictionary<string, ImportContext>();

            BackgroundWorker bw = new BackgroundWorker();
            bw.WorkerReportsProgress = true; // useless here for now
            UpdateStatusBarMessage("Analyzing PE File " + currentPe.Filepath);

            bw.DoWork += (sender, e) => { ProcessPe(newTreeContexts, currentPe); };


            bw.RunWorkerCompleted += (sender, e) =>
            {
                List<BacklogImport> peProcessingBacklog = new List<BacklogImport>();

                // Important !
                // 
                // This handler is executed in the STA (Single Thread Application)
                // which is authorized to manipulate UI elements. The BackgroundWorker is not.
                //

                foreach (ImportContext newTreeContext in newTreeContexts.Values)
                {
                    ModuleTreeViewItem childTreeNode = new ModuleTreeViewItem(rootNode);
                    DependencyNodeContext childTreeNodeContext = new DependencyNodeContext();
                    childTreeNodeContext.IsDummy = false;

                    string moduleName = newTreeContext.ModuleName;
                    string moduleFilePath = newTreeContext.PeFilePath;
                    ModuleCacheKey moduleKey = new ModuleCacheKey(newTreeContext);

                    // Newly seen modules
                    if (!_processedModulesCache.ContainsKey(moduleKey))
                    {
                        // Missing module "found"
                        if ((newTreeContext.PeFilePath == null) || !NativeFile.Exists(newTreeContext.PeFilePath))
                        {
                            if (newTreeContext.IsApiSet)
                            {
                                _processedModulesCache[moduleKey] =
                                    new ApiSetNotFoundModuleInfo(moduleName, newTreeContext.ApiSetModuleName);
                            }
                            else
                            {
                                _processedModulesCache[moduleKey] = new NotFoundModuleInfo(moduleName);
                            }
                        }
                        else
                        {
                            if (newTreeContext.IsApiSet)
                            {
                                var apiSetContractModule = new DisplayModuleInfo(
                                    newTreeContext.ApiSetModuleName,
                                    newTreeContext.PeProperties,
                                    newTreeContext.ModuleLocation,
                                    newTreeContext.Flags);
                                var newModule =
                                    new ApiSetModuleInfo(newTreeContext.ModuleName, ref apiSetContractModule);

                                _processedModulesCache[moduleKey] = newModule;

                                if (settingTreeBehaviour == TreeBuildingBehaviour.DependencyTreeBehaviour.Recursive)
                                {
                                    peProcessingBacklog.Add(new BacklogImport(childTreeNode,
                                        apiSetContractModule.ModuleName));
                                }
                            }
                            else
                            {
                                var newModule = new DisplayModuleInfo(
                                    newTreeContext.ModuleName,
                                    newTreeContext.PeProperties,
                                    newTreeContext.ModuleLocation,
                                    newTreeContext.Flags);
                                _processedModulesCache[moduleKey] = newModule;

                                switch (settingTreeBehaviour)
                                {
                                    case TreeBuildingBehaviour.DependencyTreeBehaviour.RecursiveOnlyOnDirectImports:
                                        if ((newTreeContext.Flags & ModuleFlag.DelayLoad) == 0)
                                        {
                                            peProcessingBacklog.Add(new BacklogImport(childTreeNode,
                                                newModule.ModuleName));
                                        }

                                        break;

                                    case TreeBuildingBehaviour.DependencyTreeBehaviour.Recursive:
                                        peProcessingBacklog.Add(new BacklogImport(childTreeNode, newModule.ModuleName));
                                        break;
                                }
                            }
                        }

                        // add it to the module list
                        ModulesList.AddModule(_processedModulesCache[moduleKey]);
                    }

                    // Since we uniquely process PE, for thoses who have already been "seen",
                    // we set a dummy entry in order to set the "[+]" icon next to the node.
                    // The dll dependencies are actually resolved on user double-click action
                    // We can't do the resolution in the same time as the tree construction since
                    // it's asynchronous (we would have to wait for all the background to finish and
                    // use another Async worker to resolve).

                    // Some dot net dlls give 0 for GetImports() but they will always have imports
                    // that can be detected using the special CLR dll processing we do. 
                    if ((newTreeContext.PeProperties != null) &&
                        (newTreeContext.PeProperties.GetImports().Count > 0 || newTreeContext.PeProperties.IsClrDll()))
                    {
                        ModuleTreeViewItem dummyEntry = new ModuleTreeViewItem();
                        DependencyNodeContext dummyContext = new DependencyNodeContext()
                        {
                            ModuleInfo = new WeakReference(new NotFoundModuleInfo("Dummy")),
                            IsDummy = true
                        };

                        dummyEntry.DataContext = dummyContext;
                        dummyEntry.Header = "@Dummy : if you see this header, it's a bug.";
                        dummyEntry.IsExpanded = false;

                        childTreeNode.Items.Add(dummyEntry);
                        childTreeNode.Expanded += ResolveDummyEntries;
                    }

                    // Add to tree view
                    childTreeNodeContext.ModuleInfo = new WeakReference(_processedModulesCache[moduleKey]);
                    childTreeNode.DataContext = childTreeNodeContext;
                    childTreeNode.Header =
                        childTreeNode.GetTreeNodeHeaderName(Properties.Settings.Default.FullPath);
                    rootNode.Items.Add(childTreeNode);
                }


                // Process next batch of dll imports only if :
                //	1. Recursive tree building has been activated
                //  2. Recursion is not hitting the max depth level
                bool doProcessNextLevel =
                    (settingTreeBehaviour != TreeBuildingBehaviour.DependencyTreeBehaviour.ChildOnly) &&
                    (recursionLevel < Properties.Settings.Default.TreeDepth) &&
                    peProcessingBacklog.Count > 0;

                if (doProcessNextLevel)
                {
                    foreach (var importNode in peProcessingBacklog)
                    {
                        ConstructDependencyTree(
                            importNode.Item1,
                            importNode.Item2,
                            settingTreeBehaviour,
                            recursionLevel + 1,
                            onRunWorkerCompleted); // warning : recursive call
                    }
                }

                UpdateStatusBarMessage(rootNode.ModuleFilePath + " Loaded successfully. Modules " +
                                       ModulesList.Items.Count);

                onRunWorkerCompleted?.Invoke();
            };

            bw.RunWorkerAsync();
        }

        private void UpdateStatusBarMessage(string msg)
        {
            if (Application.Current is App app)
            {
                app.StatusBarMessage = msg;
            }
        }

        /// <summary>
        /// Resolve imports when the user expand the node.
        /// </summary>
        private void ResolveDummyEntries(object sender, RoutedEventArgs e)
        {
            ResolveDummyEntries(e.OriginalSource as ModuleTreeViewItem);
        }

        private void ResolveDummyEntries(ModuleTreeViewItem moduleTreeViewItem, bool force = false)
        {
            if (moduleTreeViewItem == null) return;
            ModuleTreeViewItem needDummyPeNode = moduleTreeViewItem;
            if (!force && needDummyPeNode!.Items.Count == 0)
            {
                return;
            }

            ModuleTreeViewItem maybeDummyNode = (ModuleTreeViewItem)needDummyPeNode.Items[0];
            DependencyNodeContext context = (DependencyNodeContext)maybeDummyNode.DataContext;

            //TODO: Improve resolution predicate
            if (!context.IsDummy)
            {
                return;
            }

            needDummyPeNode.Items.Clear();
            string filepath = needDummyPeNode.ModuleFilePath;
            TreeBuildingBehaviour.DependencyTreeBehaviour settingTreeBehaviour =
                TreeBuildingBehaviour.GetGlobalBehaviour();
            ConstructDependencyTree(needDummyPeNode, filepath, settingTreeBehaviour);
        }

        #endregion TreeConstruction

        #region Commands

        private void OnModuleViewSelectedItemChanged(object sender, RoutedEventArgs e)
        {
            DisplayModuleInfo selectedModule = (sender as DependencyModuleList).SelectedItem as DisplayModuleInfo;

            // Selected Pe has not been found on disk
            if (selectedModule == null)
                return;

            // Display module as root (since we can't know which parent it's attached to)
            UpdateImportExportLists(selectedModule, null);
        }

        private void OnTreeViewSelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (DllTreeView.SelectedItem == null)
            {
                UpdateImportExportLists(null, null);
                return;
            }

            DependencyNodeContext childTreeContext =
                ((DependencyNodeContext)(DllTreeView.SelectedItem as ModuleTreeViewItem).DataContext);
            DisplayModuleInfo selectedModule = childTreeContext.ModuleInfo.Target as DisplayModuleInfo;
            if (selectedModule == null)
            {
                return;
            }

            // Selected Pe has not been found on disk : unvalidate current module
            selectedModule.HasErrors = !NativeFile.Exists(selectedModule.Filepath);
            if (selectedModule.HasErrors)
            {
                // TODO : do a proper refresh instead of asking the user to do it
                MessageBox.Show(String.Format(
                    "We could not find {0:s} file on the disk anymore, please fix this problem and refresh the window via F5",
                    selectedModule.Filepath));
            }

            // Root Item : no parent
            ModuleTreeViewItem treeRootItem = DllTreeView.Items[0] as ModuleTreeViewItem;
            ModuleTreeViewItem selectedItem = DllTreeView.SelectedItem as ModuleTreeViewItem;
            if (selectedItem == treeRootItem)
            {
                // Selected Pe has not been found on disk : unvalidate current module
                if (selectedModule.HasErrors)
                {
                    UpdateImportExportLists(null, null);
                }
                else
                {
                    selectedModule.HasErrors = false;
                    UpdateImportExportLists(selectedModule, null);
                }

                return;
            }

            // Tree Item
            DisplayModuleInfo parentModule = selectedItem.ParentModule.ModuleInfo;
            UpdateImportExportLists(selectedModule, parentModule);
        }

        private void UpdateImportExportLists(DisplayModuleInfo selectedModule, DisplayModuleInfo parent)
        {
            if (selectedModule == null)
            {
                ImportList.Items.Clear();
                ExportList.Items.Clear();
            }
            else
            {
                if (parent == null) // root module
                {
                    ImportList.SetRootImports(selectedModule.Imports, _symPrv, this);
                }
                else
                {
                    // Imports from the same dll are not necessarly sequential (see: HDDGuru\RawCopy.exe)
                    var machingImports = parent.Imports.FindAll(imp => imp.Name == selectedModule._Name);
                    ImportList.SetImports(selectedModule.Filepath, selectedModule.Exports, machingImports, _symPrv,
                        this);
                }

                ImportList.ResetAutoSortProperty();

                ExportList.SetExports(selectedModule.Exports, _symPrv);
                ExportList.ResetAutoSortProperty();
            }
        }

        public PE LoadImport(string moduleName, DisplayModuleInfo currentModule = null, bool delayLoad = false)
        {
            if (currentModule == null)
            {
                currentModule = _selectedModule;
            }

            Tuple<ModuleSearchStrategy, PE> resolvedModule = BinaryCache.ResolveModule(
                _pe,
                moduleName,
                _sxsEntriesCache,
                CustomSearchFolders,
                WorkingDirectory
            );

            string moduleFilepath = (resolvedModule.Item2 != null) ? resolvedModule.Item2.Filepath : null;

            // Not found module, returning PE without update module list
            if (moduleFilepath == null)
            {
                return resolvedModule.Item2;
            }

            ModuleFlag moduleFlags = ModuleFlag.NoFlag;
            if (delayLoad)
                moduleFlags |= ModuleFlag.DelayLoad;
            if (resolvedModule.Item1 == ModuleSearchStrategy.ApiSetSchema)
                moduleFlags |= ModuleFlag.ApiSet;

            ModuleCacheKey moduleKey = new ModuleCacheKey(moduleName, moduleFilepath, moduleFlags);
            if (!_processedModulesCache.ContainsKey(moduleKey))
            {
                DisplayModuleInfo newModule;

                // apiset resolution are a bit trickier
                if (resolvedModule.Item1 == ModuleSearchStrategy.ApiSetSchema)
                {
                    var apiSetContractModule = new DisplayModuleInfo(
                        BinaryCache.LookupApiSetLibrary(moduleName),
                        resolvedModule.Item2,
                        resolvedModule.Item1,
                        moduleFlags
                    );
                    newModule = new ApiSetModuleInfo(moduleName, ref apiSetContractModule);
                }
                else
                {
                    newModule = new DisplayModuleInfo(
                        moduleName,
                        resolvedModule.Item2,
                        resolvedModule.Item1,
                        moduleFlags
                    );
                }

                _processedModulesCache[moduleKey] = newModule;

                // add it to the module list
                ModulesList.AddModule(_processedModulesCache[moduleKey]);
            }

            return resolvedModule.Item2;
        }


        /// <summary>
        /// Reentrant version of Collapse/Expand Node
        /// </summary>
        /// <param name="item"></param>
        /// <param name="expandNode"></param>
        private void CollapseOrExpandAllNodes(ModuleTreeViewItem item, bool expandNode)
        {
            item.IsExpanded = expandNode;
            foreach (ModuleTreeViewItem childItem in item.Items)
            {
                CollapseOrExpandAllNodes(childItem, expandNode);
            }
        }

        private void ExpandAllNodes_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            // Expanding all nodes tends to slow down the application (massive allocations for node DataContext)
            // TODO : Reduce memory pressure by storing tree nodes data context in a HashSet and find an async trick
            // to improve the command responsiveness.
            TreeView treeNode = sender as TreeView;
            if (treeNode == null) return;
            CollapseOrExpandAllNodes((treeNode.Items[0] as ModuleTreeViewItem), true);
        }

        private void CollapseAllNodes_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            TreeView treeNode = sender as TreeView;
            if (treeNode == null) return;
            CollapseOrExpandAllNodes((treeNode.Items[0] as ModuleTreeViewItem), false);
        }

        private void DoFindModuleInList_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ModuleTreeViewItem source = e.Source as ModuleTreeViewItem;
            if (source == null) return;
            String selectedModuleName = source.GetTreeNodeHeaderName(Properties.Settings.Default.FullPath);

            foreach (DisplayModuleInfo item in ModulesList.Items)
            {
                if (item.ModuleName == selectedModuleName)
                {
                    ModulesList.SelectedItem = item;
                    ModulesList.ScrollIntoView(item);
                    return;
                }
            }
        }

        private void CopyFilePath_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ModuleTreeViewItem source = e.Source as ModuleTreeViewItem;
            if (source == null)
                return;
            String selectedModuleName = source.GetTreeNodeHeaderName(Properties.Settings.Default.FullPath);
            Clipboard.SetText(selectedModuleName);
        }

        private void CopyAllFileRecursionNoSystem_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            CopyAllFileRecursion(sender, e, new List<string>()
            {
                @"C:\Windows\"
            });
        }

        private void CopyAllFileRecursion_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            CopyAllFileRecursion(sender, e, new List<string>());
        }

        private void CopyAllFileRecursion(object sender, ExecutedRoutedEventArgs e,
            List<string> excludeDirectory)
        {
            ModuleTreeViewItem moduleTreeView = e.Source as ModuleTreeViewItem;
            if (moduleTreeView == null)
            {
                TreeView treeNode = sender as TreeView;
                if (treeNode == null || treeNode.Items.Count <= 0) return;
                moduleTreeView = treeNode.Items[0] as ModuleTreeViewItem;
            }

            if (moduleTreeView == null || moduleTreeView.ModuleInfo == null ||
                moduleTreeView.ModuleInfo.Imports == null) return;

            UpdateStatusBarMessage($"Prepare for Copyed Files,Please waiting a moment ...");

            Dictionary<string, ImportContext> newTreeContexts = new();
            SxsEntries sxsEntries = new();
            Dictionary<string, DisplayModuleInfo> modulesCache = new();
            List<DisplayModuleInfo> moduleList = new();

            var sc = new StringCollection();
            sc.Add(moduleTreeView.ModuleFilePath);

            foreach (ModuleTreeViewItem treeItem in moduleTreeView.Items.Cast<ModuleTreeViewItem>())
            {
                var dir = Path.GetDirectoryName(treeItem.ModuleFilePath)!;
                if (!excludeDirectory.Any(d => dir.StartsWith(d, StringComparison.OrdinalIgnoreCase)))
                {
                    sc.Add(treeItem.ModuleFilePath);
                }

                PE newPe = LoadPe(treeItem.ModuleFilePath);
                //ProcessPe(RootFolder, _pe, newTreeContexts, newPe, sxsEntries, CustomSearchFolders, WorkingDirectory);
                ConstructDependencyTreeV2(
                    RootFolder,
                    _pe,
                    newPe,
                    newTreeContexts,
                    sxsEntries,
                    CustomSearchFolders,
                    WorkingDirectory,
                    modulesCache,
                    moduleList,
                    TreeBuildingBehaviour.DependencyTreeBehaviour.RecursiveOnlyOnDirectImports,
                    -8);
            }

            foreach (var newTreeContext in newTreeContexts)
            {
                if (newTreeContext.Value.ModuleLocation == ModuleSearchStrategy.NOT_FOUND) continue;
                var dir = Path.GetDirectoryName(newTreeContext.Value.PeFilePath)!;
                if (excludeDirectory.Any(d => dir.StartsWith(d, StringComparison.OrdinalIgnoreCase))) continue;
                sc.Add(newTreeContext.Value.PeFilePath);
            }

            UpdateStatusBarMessage($"Copyed {sc.Count} Files to Clipboard");
            Clipboard.SetFileDropList(sc);
        }

        private void GetRecursionModules(
            ModuleTreeViewItem root,
            List<ModuleTreeViewItem> outputs,
            Dictionary<string, ImportContext> newTreeContexts)
        {
            if (root == null || root.Items.Count <= 0) return;
            foreach (var item in root.Items.Cast<ModuleTreeViewItem>())
            {
                outputs.Add(item);
                GetRecursionModules(item, outputs, newTreeContexts);
            }
        }

        private void CopyAllFile_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ModuleTreeViewItem moduleTreeView = e.Source as ModuleTreeViewItem;
            if (moduleTreeView == null)
            {
                TreeView treeNode = sender as TreeView;
                if (treeNode == null || treeNode.Items.Count <= 0) return;
                moduleTreeView = treeNode.Items[0] as ModuleTreeViewItem;
            }

            if (moduleTreeView == null || moduleTreeView.ModuleInfo == null ||
                moduleTreeView.ModuleInfo.Imports == null) return;
            var sc = new StringCollection();
            foreach (ModuleTreeViewItem treeItem in moduleTreeView.Items.Cast<ModuleTreeViewItem>())
            {
                sc.Add(treeItem.ModuleFilePath);
            }

            UpdateStatusBarMessage($"Copyed {sc.Count} Files to Clipboard");
            Clipboard.SetFileDropList(sc);
        }

        private void CopyFile_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ModuleTreeViewItem source = e.Source as ModuleTreeViewItem;
            if (source == null) return;
            if (!File.Exists(source.ModuleFilePath)) return;
            Clipboard.SetFileDropList(new StringCollection()
            {
                source.ModuleFilePath
            });
            UpdateStatusBarMessage($"{source.ModuleFilePath} Copyed to Clipboard");
        }

        private void OpenInExplorer_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ModuleTreeViewItem source = e.Source as ModuleTreeViewItem;
            if (source == null)
                return;

            String selectedModuleName = source.GetTreeNodeHeaderName(Properties.Settings.Default.FullPath);
            String commandParameter = "/select,\"" + selectedModuleName + "\"";

            Process.Start("explorer.exe", commandParameter);
        }

        private void ExpandAllParentNode(ModuleTreeViewItem item)
        {
            if (item != null)
            {
                ExpandAllParentNode(item.Parent as ModuleTreeViewItem);
                item.IsExpanded = true;
            }
        }

        /// <summary>
        /// Reentrant version of Collapse/Expand Node
        /// </summary>
        /// <param name="item"></param>
        /// <param name="ExpandNode"></param>
        private ModuleTreeViewItem FindModuleInTree(ModuleTreeViewItem item, DisplayModuleInfo module,
            bool highlight = false)
        {
            if (item.GetTreeNodeHeaderName(Properties.Settings.Default.FullPath) == module.ModuleName)
            {
                if (highlight)
                {
                    ExpandAllParentNode(item.Parent as ModuleTreeViewItem);
                    item.IsSelected = true;
                    item.BringIntoView();
                    item.Focus();
                }

                return item;
            }

            // BFS style search -> return the first matching node with the lowest "depth"
            foreach (ModuleTreeViewItem childItem in item.Items)
            {
                if (childItem.GetTreeNodeHeaderName(Properties.Settings.Default.FullPath) ==
                    module.ModuleName)
                {
                    if (highlight)
                    {
                        ExpandAllParentNode(item);
                        childItem.IsSelected = true;
                        childItem.BringIntoView();
                        childItem.Focus();
                    }

                    return item;
                }
            }

            foreach (ModuleTreeViewItem childItem in item.Items)
            {
                ModuleTreeViewItem matchingItem = FindModuleInTree(childItem, module, highlight);

                // early exit as soon as we find a matching node
                if (matchingItem != null)
                    return matchingItem;
            }

            return null;
        }


        public RelayCommand DoFindModuleInTree
        {
            get
            {
                return new RelayCommand((param) =>
                {
                    DisplayModuleInfo selectedModule = (param as DisplayModuleInfo);
                    ModuleTreeViewItem treeRootItem = DllTreeView.Items[0] as ModuleTreeViewItem;

                    FindModuleInTree(treeRootItem, selectedModule, true);
                });
            }
        }

        public RelayCommand ConfigureSearchOrderCommand
        {
            get
            {
                return new RelayCommand((param) =>
                {
                    ModuleSearchOrder modalWindow = new ModuleSearchOrder(_processedModulesCache);
                    modalWindow.ShowDialog();
                });
            }
        }

        #endregion // Commands 
    }
}