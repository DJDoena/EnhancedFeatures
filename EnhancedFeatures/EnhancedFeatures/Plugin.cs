using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Text;
using System.Web;
using System.Windows.Forms;
using DoenaSoft.DVDProfiler.DVDProfilerHelper;
using DoenaSoft.DVDProfiler.EnhancedFeatures.Resources;
using Invelos.DVDProfilerPlugin;

namespace DoenaSoft.DVDProfiler.EnhancedFeatures
{
    [Guid(ClassGuid.ClassID), ComVisible(true)]
    public class Plugin : IDVDProfilerPlugin, IDVDProfilerPluginInfo, IDVDProfilerDataAwarePlugin
    {
        public const Byte FeatureCount = 40;

        private readonly String SettingsFile;

        private readonly String ErrorFile;

        private readonly String ApplicationPath;

        internal IDVDProfilerAPI Api { get; private set; }

        internal Settings Settings { get; private set; }

        private Boolean DatabaseRestoreRunning = false;

        internal Boolean IsRemoteAccess { get; private set; }

        #region MenuTokens

        private String DvdMenuToken = "";

        private const Int32 DvdMenuId = 1;

        private String CollectionExportMenuToken = "";

        private const Int32 CollectionExportMenuId = 21;

        private String CollectionImportMenuToken = "";

        private const Int32 CollectionImportMenuId = 22;

        private String CollectionExportToCsvMenuToken = "";

        private const Int32 CollectionExportToCsvMenuId = 23;

        private String CollectionFlaggedExportMenuToken = "";

        private const Int32 CollectionFlaggedExportMenuId = 31;

        private String CollectionFlaggedImportMenuToken = "";

        private const Int32 CollectionFlaggedImportMenuId = 32;

        private String CollectionFlaggedExportToCsvMenuToken = "";

        private const Int32 CollectionFlaggedExportToCsvMenuId = 33;

        private String ToolsOptionsMenuToken = "";

        private const Int32 ToolsOptionsMenuId = 41;

        private String ToolsExportOptionsMenuToken = "";

        private const Int32 ToolsExportOptionsMenuId = 42;

        private String ToolsImportOptionsMenuToken = "";

        private const Int32 ToolsImportOptionsMenuId = 43;

        #endregion

        private readonly Dictionary<String, String> FilterTokens;

        public Plugin()
        {
            FilterTokens = new Dictionary<String, String>();
            ApplicationPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Doena Soft\EnhancedFeatures\";
            SettingsFile = ApplicationPath + "EnhancedFeatures.xml";
            ErrorFile = Environment.GetEnvironmentVariable("TEMP") + @"\EnhancedFeaturesCrash.xml";
        }

        #region I.. Members

        #region IDVDProfilerPlugin Members

        public void Load(IDVDProfilerAPI api)
        {
            //System.Diagnostics.Debugger.Launch();

            Api = api;

            if (Directory.Exists(ApplicationPath) == false)
            {
                Directory.CreateDirectory(ApplicationPath);
            }

            if (File.Exists(SettingsFile))
            {
                try
                {
                    Settings = Serializer<Settings>.Deserialize(SettingsFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(String.Format(MessageBoxTexts.FileCantBeRead, SettingsFile, ex.Message)
                        , MessageBoxTexts.ErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            EnsureSettingsAndSetUILanguage();

            SetIsRemoteAccess();

            RegisterForEvents();

            RegisterMenuItems();

            RegisterCustomFields();
        }

        public void Unload()
        {
            Api.UnregisterMenuItem(DvdMenuToken);

            Api.UnregisterMenuItem(CollectionExportMenuToken);
            Api.UnregisterMenuItem(CollectionImportMenuToken);
            Api.UnregisterMenuItem(CollectionExportToCsvMenuToken);

            Api.UnregisterMenuItem(CollectionFlaggedExportMenuToken);
            Api.UnregisterMenuItem(CollectionFlaggedImportMenuToken);
            Api.UnregisterMenuItem(CollectionFlaggedExportToCsvMenuToken);

            //Api.UnregisterMenuItem(PluginInfoToolsMenuToken);
            Api.UnregisterMenuItem(ToolsOptionsMenuToken);
            Api.UnregisterMenuItem(ToolsExportOptionsMenuToken);
            Api.UnregisterMenuItem(ToolsImportOptionsMenuToken);

            try
            {
                Serializer<Settings>.Serialize(SettingsFile, Settings);
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format(MessageBoxTexts.FileCantBeWritten, SettingsFile, ex.Message)
                    , MessageBoxTexts.ErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Api = null;
        }

        public void HandleEvent(Int32 EventType
            , Object EventData)
        {
            try
            {
                switch (EventType)
                {
                    case (PluginConstants.EVENTID_CustomMenuClick):
                        {
                            HandleMenuClick((Int32)EventData);

                            break;
                        }
                    case (PluginConstants.EVENTID_RestoreStarting):
                        {
                            DatabaseRestoreRunning = true;

                            RegisterCustomFields(false);

                            break;
                        }
                    case (PluginConstants.EVENTID_DatabaseOpened):
                    case (PluginConstants.EVENTID_RestoreFinished):
                    case (PluginConstants.EVENTID_RestoreCancelled):
                        {
                            DatabaseRestoreRunning = false;

                            RegisterCustomFields();

                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    MessageBox.Show(String.Format(MessageBoxTexts.CriticalError, ex.Message, ErrorFile)
                        , MessageBoxTexts.CriticalErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    if (File.Exists(ErrorFile))
                    {
                        File.Delete(ErrorFile);
                    }

                    LogException(ex);
                }
                catch (Exception inEx)
                {
                    MessageBox.Show(String.Format(MessageBoxTexts.FileCantBeWritten, ErrorFile, inEx.Message)
                        , MessageBoxTexts.ErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        #endregion

        #region IDVDProfilerPluginInfo Members

        public String GetName()
            => (Texts.PluginName);

        public String GetDescription()
            => (Texts.PluginDescription);

        public String GetAuthorName()
            => ("Doena Soft.");

        public String GetAuthorWebsite()
            => (Texts.PluginUrl);

        public Int32 GetPluginAPIVersion()
            => (PluginConstants.API_VERSION);

        public Int32 GetVersionMajor()
        {
            Version version = Assembly.GetAssembly(GetType()).GetName().Version;

            Int32 major = version.Major;

            return (major);
        }

        public Int32 GetVersionMinor()
        {
            Version version = Assembly.GetAssembly(GetType()).GetName().Version;

            Int32 minor = version.Minor * 100 + version.Build * 10 + version.Revision;

            return (minor);
        }

        #endregion

        #region IDVDProfilerDataAwarePlugin

        public Boolean ExportsCustomDataXML()
            => ((Settings.DefaultValues.ExportToCollectionXml) && (DatabaseRestoreRunning == false));

        public String GetCustomDataXMLForDVD(IDVDInfo SourceDVD)
        {
            if (Settings.DefaultValues.ExportToCollectionXml == false)
            {
                return (String.Empty);
            }
            else if (DatabaseRestoreRunning)
            {
                return (String.Empty);
            }

            FeatureManager featureManager = new FeatureManager(SourceDVD);

            Boolean[] hasFeatures = new Boolean[FeatureCount];

            for (Byte featureIndex = 1; featureIndex <= FeatureCount; featureIndex++)
            {
                hasFeatures[featureIndex - 1] = featureManager.GetFeature(featureIndex);
            }

            String xml = String.Empty;

            if (hasFeatures.Any(feature => feature))
            {
                DefaultValues dv = Settings.DefaultValues;

                StringBuilder sb = new StringBuilder("<EnhancedFeatures>");

                for (Byte featureIndex = 1; featureIndex <= FeatureCount; featureIndex++)
                {
                    if (hasFeatures[featureIndex - 1])
                    {
                        AddTag(sb, featureIndex, dv.FeatureLabels[featureIndex]);
                    }
                }

                sb.Append("</EnhancedFeatures>");

                xml = sb.ToString();
            }

            return (xml);
        }

        public String GetHTMLForDPVarsFunctionSection()
            => (String.Empty);

        public String GetHTMLForDPVarsDataSection(IDVDInfo SourceDVD
            , IDVDInfo CompareDVD)
            => (String.Empty);

        public String GetHTMLForTag(String TagName
            , String FullTag
            , IDVDInfo SourceDVD
            , IDVDInfo CompareDVD
            , out Boolean Handled)
        {
            if (String.IsNullOrEmpty(TagName))
            {
                Handled = false;

                return (null);
            }
            else if (TagName.StartsWith(Constants.HtmlPrefix + ".") == false)
            {
                Handled = false;

                return (null);
            }

            FeatureManager featureManager = new FeatureManager(SourceDVD);

            Handled = false;

            String text = null;

            DefaultValues dv = Settings.DefaultValues;

            Int32 prefixLength = Constants.HtmlPrefix.Length + 1 + Constants.Feature.Length;

            if (TagName.Length > prefixLength)
            {
                TagName = TagName.Substring(prefixLength);

                if (TagName.EndsWith(Constants.LabelSuffix) == false)
                {
                    Byte featureIndex;
                    if (Byte.TryParse(TagName, out featureIndex))
                    {
                        Boolean value = featureManager.GetFeature(featureIndex);

                        text = value ? "X" : String.Empty;

                        Handled = true;
                    }
                }
                else
                {
                    TagName = TagName.Substring(0, TagName.Length - Constants.LabelSuffix.Length);

                    Byte featureIndex;
                    if (Byte.TryParse(TagName, out featureIndex))
                    {
                        text = HtmlEncode(dv.FeatureLabels[featureIndex]);

                        Handled = true;
                    }
                }
            }

            return (text);
        }

        public Object GetCustomHTMLTagNames()
        {
            String[] tags = new String[FeatureCount * 2];

            for (Byte featureIndex = 1; featureIndex <= FeatureCount; featureIndex++)
            {
                tags[featureIndex - 1] = $"{Constants.HtmlPrefix}.{Constants.Feature}{featureIndex}";

                tags[featureIndex - 1 + FeatureCount] = $"{Constants.HtmlPrefix}.{Constants.Feature}{featureIndex}{Constants.LabelSuffix}";
            }

            return (tags);
        }

        public Object GetCustomHTMLParamsForTag(String TagName)
            => (null);

        public Boolean FilterFieldMatch(String FieldFilterToken
            , Int32 ComparisonTypeIndex
            , Object ComparisonValue
            , IDVDInfo TestDVD)
        {
            if ((FieldFilterToken == null) || (ComparisonValue == null))
            {
                return (false);
            }

            String fieldName;
            if ((FilterTokens.TryGetValue(FieldFilterToken, out fieldName)) && (fieldName.Length > Constants.Feature.Length))
            {
                fieldName = fieldName.Substring(Constants.Feature.Length);

                Byte featureIndex;
                if (Byte.TryParse(fieldName, out featureIndex))
                {
                    FeatureManager featureManager = new FeatureManager(TestDVD);

                    Boolean actual = featureManager.GetFeature(featureIndex);

                    Boolean expected = (Boolean)ComparisonValue;

                    return (actual == expected);
                }
            }

            return (false);
        }

        #endregion

        #endregion

        #region RegisterCustomFields

        internal void RegisterCustomFields(Boolean rebuildFilters = true)
        {
            try
            {
                UnregisterCustomFilterField(rebuildFilters);

                #region Schema

                using (Stream stream = typeof(EnhancedFeatures).Assembly.GetManifestResourceStream("DoenaSoft.DVDProfiler.EnhancedFeatures.EnhancedFeatures.xsd"))
                {
                    using (StreamReader sr = new StreamReader(stream))
                    {
                        String xsd = sr.ReadToEnd();

                        Api.SetGlobalSetting(Constants.FieldDomain, "EnhancedFeaturesSchema", xsd, Constants.ReadKey, InternalConstants.WriteKey);
                    }
                }

                #endregion

                DefaultValues dv = Settings.DefaultValues;

                //System.Diagnostics.Debugger.Launch();

                for (Byte featureIndex = 1; featureIndex <= FeatureCount; featureIndex++)
                {
                    RegisterCustomField($"{Constants.Feature}{featureIndex}", dv.FeatureLabels[featureIndex], rebuildFilters, featureIndex);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    MessageBox.Show(String.Format(MessageBoxTexts.CriticalError, ex.Message, ErrorFile)
                        , MessageBoxTexts.CriticalErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Stop);

                    if (File.Exists(ErrorFile))
                    {
                        File.Delete(ErrorFile);
                    }

                    LogException(ex);
                }
                catch (Exception inEx)
                {
                    MessageBox.Show(String.Format(MessageBoxTexts.FileCantBeWritten, ErrorFile, inEx.Message)
                        , MessageBoxTexts.ErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void UnregisterCustomFilterField(Boolean rebuildFilters)
        {
            //System.Diagnostics.Debugger.Launch();

            if (rebuildFilters)
            {
                foreach (String token in FilterTokens.Keys)
                {
                    try
                    {
                        Api.RemoveCustomFilterField(token);
                    }
                    catch (COMException)
                    { }
                }

                FilterTokens.Clear();
            }
        }

        private void RegisterCustomField(String fieldName
            , String displayName
            , Boolean rebuildFilters
            , Byte featureIndex)
        {
            Api.CreateCustomDVDField(Constants.FieldDomain, fieldName, PluginConstants.FIELD_TYPE_BOOL, Constants.ReadKey, InternalConstants.WriteKey);

            Api.SetCustomDVDFieldStorage(Constants.FieldDomain, fieldName, InternalConstants.WriteKey, true, false);

            if (featureIndex <= Settings.DefaultValues.FilterCount)
            {
                RegisterCustomFilterField(fieldName, displayName, rebuildFilters);
            }
        }

        private void RegisterCustomFilterField(String fieldName
            , String displayName
            , Boolean rebuildFilters)
        {
            if (rebuildFilters)
            {
                if (displayName == null)
                {
                    ResourceManager rm = Texts.ResourceManager;

                    displayName = rm.GetString(fieldName);
                }

                //System.Diagnostics.Debugger.Launch();

                String token = Api.SetCustomFieldFilterableA(displayName, PluginConstants.FILTER_INPUT_CHECKBOX, null, null);

                if (token != null)
                {
                    FilterTokens.Add(token, fieldName);
                }
                else
                {
                    System.Diagnostics.Debug.Fail("No valid token for search!");
                }
            }
        }

        #endregion

        private void SetIsRemoteAccess()
        {
            String name;
            Boolean isRemote;
            String localPath;
            Api.GetCurrentDatabaseInformation(out name, out isRemote, out localPath);

            //System.Diagnostics.Debugger.Launch();

            IsRemoteAccess = isRemote;
        }

        private void RegisterForEvents()
        {
            Api.RegisterForEvent(PluginConstants.EVENTID_DatabaseOpened);

            Api.RegisterForEvent(PluginConstants.EVENTID_RestoreStarting);
            Api.RegisterForEvent(PluginConstants.EVENTID_RestoreFinished);
            Api.RegisterForEvent(PluginConstants.EVENTID_RestoreCancelled);
        }

        private void RegisterMenuItems()
        {
            DvdMenuToken = Api.RegisterMenuItemA(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
                , "DVD", Texts.EF, DvdMenuId, "", PluginConstants.SHORTCUT_KEY_A + 5, PluginConstants.SHORTCUT_MOD_Ctrl + PluginConstants.SHORTCUT_MOD_Shift, false);

            CollectionExportMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
                , @"Collection\" + Texts.EF, Texts.ExportToXml, CollectionExportMenuId);

            if (IsRemoteAccess == false)
            {
                CollectionImportMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
                    , @"Collection\" + Texts.EF, Texts.ImportFromXml, CollectionImportMenuId);
            }

            CollectionExportToCsvMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
                , @"Collection\" + Texts.EF, Texts.ExportToExcel, CollectionExportToCsvMenuId);

            CollectionFlaggedExportMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
                , @"Collection\Flagged\" + Texts.EF, Texts.ExportToXml, CollectionFlaggedExportMenuId);

            if (IsRemoteAccess == false)
            {
                CollectionFlaggedImportMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
                   , @"Collection\Flagged\" + Texts.EF, Texts.ImportFromXml, CollectionFlaggedImportMenuId);
            }

            CollectionFlaggedExportToCsvMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
                , @"Collection\Flagged\" + Texts.EF, Texts.ExportToExcel, CollectionFlaggedExportToCsvMenuId);

            ToolsOptionsMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
               , @"Tools\" + Texts.EF, Texts.Options, ToolsOptionsMenuId);
            ToolsExportOptionsMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
                , @"Tools\" + Texts.EF, Texts.ExportOptions, ToolsExportOptionsMenuId);
            ToolsImportOptionsMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
                , @"Tools\" + Texts.EF, Texts.ImportOptions, ToolsImportOptionsMenuId);

            //PluginInfoToolsMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
            //    , "Tools", "Show Plugin FieldAccess", PluginInfoToolsMenuId);
        }

        private static void AddTag(StringBuilder sb
            , Byte featureIndex
            , String displayName)
        {
            sb.Append("<");
            sb.Append(Constants.Feature);
            sb.Append(" Index =\"");
            sb.Append(featureIndex);
            sb.Append("\"");

            String base64;
            displayName = XmlConvertHelper.GetWindows1252Text(displayName, out base64);

            sb.Append(" DisplayName=\"");
            sb.Append(displayName);
            sb.Append("\"");

            if (base64 != null)
            {
                sb.Append(" Base64DisplayName=\"");
                sb.Append(base64);
                sb.Append("\"");
            }

            sb.Append(">true</");
            sb.Append(Constants.Feature);
            sb.Append(">");
        }

        private void EnsureSettingsAndSetUILanguage()
        {
            Texts.Culture = DefaultValues.GetUILanguage();

            CultureInfo uiLanguage = EnsureSettings();

            Texts.Culture = uiLanguage;

            MessageBoxTexts.Culture = uiLanguage;
        }

        private CultureInfo EnsureSettings()
        {
            if (Settings == null)
            {
                Settings = new Settings();
            }

            if (Settings.DefaultValues == null)
            {
                Settings.DefaultValues = new DefaultValues();
            }

            return (Settings.DefaultValues.UiLanguage);
        }

        private static String HtmlEncode(String decoded)
        {
            String encoded = String.Join("", decoded.ToCharArray().Select(c =>
                    {
                        Int32 number = c;

                        String newChar = (number > 127) ? ("&#" + number.ToString() + ";") : (HttpUtility.HtmlEncode(c.ToString()));

                        return (newChar);
                    }
                ).ToArray());

            return (encoded);
        }

        private void HandleMenuClick(Int32 MenuEventID)
        {
            try
            {
                switch (MenuEventID)
                {
                    case (DvdMenuId):
                        {
                            OpenEditor();

                            break;
                        }
                    case (CollectionExportMenuId):
                        {
                            XmlManager xmlManager = new XmlManager(this);

                            xmlManager.Export(true);

                            break;
                        }
                    case (CollectionImportMenuId):
                        {
                            XmlManager xmlManager = new XmlManager(this);

                            xmlManager.Import(true);

                            break;
                        }
                    case (CollectionExportToCsvMenuId):
                        {
                            CsvManager csvManager = new CsvManager(this);

                            csvManager.Export(true);

                            break;
                        }
                    case (CollectionFlaggedExportMenuId):
                        {
                            XmlManager xmlManager = new XmlManager(this);

                            xmlManager.Export(false);

                            break;
                        }
                    case (CollectionFlaggedImportMenuId):
                        {
                            XmlManager xmlManager = new XmlManager(this);

                            xmlManager.Import(false);

                            break;
                        }
                    case (CollectionFlaggedExportToCsvMenuId):
                        {
                            CsvManager csvManager = new CsvManager(this);

                            csvManager.Export(false);

                            break;
                        }
                    case (ToolsOptionsMenuId):
                        {
                            OpenSettings();

                            break;
                        }
                    case (ToolsExportOptionsMenuId):
                        {
                            ExportOptions();

                            break;
                        }
                    case (ToolsImportOptionsMenuId):
                        {
                            ImportOptions();

                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    MessageBox.Show(String.Format(MessageBoxTexts.CriticalError, ex.Message, ErrorFile)
                        , MessageBoxTexts.CriticalErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Stop);

                    if (File.Exists(ErrorFile))
                    {
                        File.Delete(ErrorFile);
                    }

                    LogException(ex);
                }
                catch (Exception inEx)
                {
                    MessageBox.Show(String.Format(MessageBoxTexts.FileCantBeWritten, ErrorFile, inEx.Message), MessageBoxTexts.ErrorHeader
                        , MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        internal void ImportOptions()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.CheckFileExists = true;
                ofd.Filter = "XML files|*.xml";
                ofd.Multiselect = false;
                ofd.RestoreDirectory = true;
                ofd.Title = Texts.LoadXmlFile;

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    DefaultValues dv = null;

                    try
                    {
                        dv = Serializer<DefaultValues>.Deserialize(ofd.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(String.Format(MessageBoxTexts.FileCantBeRead, ofd.FileName, ex.Message)
                           , MessageBoxTexts.ErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    if (dv != null)
                    {
                        Settings.DefaultValues = dv;
                        Texts.Culture = dv.UiLanguage;
                        MessageBoxTexts.Culture = dv.UiLanguage;
                        MessageBox.Show(MessageBoxTexts.Done, MessageBoxTexts.InformationHeader, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        internal void ExportOptions()
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.AddExtension = true;
                sfd.DefaultExt = ".xml";
                sfd.Filter = "XML files|*.xml";
                sfd.OverwritePrompt = true;
                sfd.RestoreDirectory = true;
                sfd.Title = Texts.SaveXmlFile;
                sfd.FileName = "EnhancedFeaturesOptions.xml";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    DefaultValues dv = Settings.DefaultValues;

                    try
                    {
                        Serializer<DefaultValues>.Serialize(sfd.FileName, dv);

                        MessageBox.Show(MessageBoxTexts.Done, MessageBoxTexts.InformationHeader, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(String.Format(MessageBoxTexts.FileCantBeWritten, sfd.FileName, ex.Message)
                            , MessageBoxTexts.ErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void OpenSettings()
        {
            using (SettingsForm form = new SettingsForm(this))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    RegisterCustomFields();
                }
            }
        }

        private void OpenEditor()
        {
            IDVDInfo profile = Api.GetDisplayedDVD();

            String profileId = profile.GetProfileID();

            if (String.IsNullOrEmpty(profileId) == false)
            {
                Api.DVDByProfileID(out profile, profileId, PluginConstants.DATASEC_AllSections, -1);

                using (MainForm form = new MainForm(this, profile))
                {
                    form.ShowDialog();
                }
            }
        }

        private void LogException(Exception ex)
        {
            ex = WrapCOMException(ex);

            ExceptionXml exceptionXml = new ExceptionXml(ex);

            Serializer<ExceptionXml>.Serialize(ErrorFile, exceptionXml);
        }

        private Exception WrapCOMException(Exception ex)
        {
            Exception returnEx = ex;

            COMException comEx = ex as COMException;

            if (comEx != null)
            {
                String lastApiError = Api.GetLastError();

                EnhancedCOMException newEx = new EnhancedCOMException(comEx, lastApiError);

                returnEx = newEx;
            }

            return (returnEx);
        }

        #region Plugin Registering

        [DllImport("user32.dll")]
        public extern static int SetParent(int child, int parent);

        [ComImport(), Guid("0002E005-0000-0000-C000-000000000046")]
        internal class StdComponentCategoriesMgr { }

        [ComRegisterFunction()]
        public static void RegisterServer(Type t)
        {
            CategoryRegistrar.ICatRegister cr = (CategoryRegistrar.ICatRegister)new StdComponentCategoriesMgr();

            Guid clsidThis = new Guid(ClassGuid.ClassID);

            Guid catid = new Guid("833F4274-5632-41DB-8FC5-BF3041CEA3F1");

            cr.RegisterClassImplCategories(ref clsidThis, 1, new[] { catid });
        }

        [ComUnregisterFunction()]
        public static void UnregisterServer(Type t)
        {
            CategoryRegistrar.ICatRegister cr = (CategoryRegistrar.ICatRegister)new StdComponentCategoriesMgr();

            Guid clsidThis = new Guid(ClassGuid.ClassID);

            Guid catid = new Guid("833F4274-5632-41DB-8FC5-BF3041CEA3F1");

            cr.UnRegisterClassImplCategories(ref clsidThis, 1, new[] { catid });
        }

        #endregion
    }
}