using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;
using DoenaSoft.DVDProfiler.EnhancedFeatures.Resources;

namespace DoenaSoft.DVDProfiler.EnhancedFeatures
{
    public partial class SettingsForm : Form
    {
        private readonly Plugin Plugin;

        public SettingsForm(Plugin plugin)
        {
            Plugin = plugin;

            InitializeComponent();

            SetSettings();

            SetLabels();

            SetComboBoxes();
        }

        private void SetSettings()
        {
            SetInvelosSettings();

            SetPluginSettings();

            DefaultValues dv = Plugin.Settings.DefaultValues;

            ExportToCollectionXmlCheckBox.Checked = dv.ExportToCollectionXml;

            FilterCountUpDown.Value = dv.FilterCount;
        }

        private void SetInvelosSettings()
        {
            DefaultValues dv = Plugin.Settings.DefaultValues;

            IdCheckBox.Checked = dv.Id;
            TitleCheckBox.Checked = dv.Title;
            EditionCheckBox.Checked = dv.Edition;
            SortTitleCheckBox.Checked = dv.SortTitle;

            SceneAccessCheckBox.Checked = dv.SceneAccess;
            PlayAllCheckBox.Checked = dv.PlayAll;
            FeatureTrailersCheckBox.Checked = dv.FeatureTrailers;
            BonusTrailerCheckBox.Checked = dv.BonusTrailers;

            FeaturettesCheckBox.Checked = dv.Featurettes;
            CommentaryCheckBox.Checked = dv.Commentary;
            DeletedScenesCheckBox.Checked = dv.DeletedScenes;
            InterviewsCheckBox.Checked = dv.Interviews;
            OuttakesBloopersCheckBox.Checked = dv.OuttakesBloopers;

            StoryboardComparisonsCheckBox.Checked = dv.StoryboardComparisons;
            GalleryCheckBox.Checked = dv.Gallery;
            ProductionNotesBiosCheckBox.Checked = dv.ProductionNotesBios;

            DVDROMContentCheckBox.Checked = dv.DVDROMContent;
            InteractiveGamesCheckBox.Checked = dv.InteractiveGames;
            MultiAngleCheckBox.Checked = dv.MultiAngle;
            MusicVideosCheckBox.Checked = dv.MusicVideos;

            THXCertifiedCheckBox.Checked = dv.THXCertified;
            ClosedCaptionedCheckBox.Checked = dv.ClosedCaptioned;

            DigitalCopyCheckBox.Checked = dv.DigitalCopy;
            PictureInPictureCheckBox.Checked = dv.PictureInPicture;
            BDLiveCheckBox.Checked = dv.BDLive;
            DBoxCheckBox.Checked = dv.DBox;
            CineChatCheckBox.Checked = dv.CineChat;
            MovieIQCheckBox.Checked = dv.MovieIQ;
        }

        private void SetPluginSettings()
        {
            DefaultValues dv = Plugin.Settings.DefaultValues;

            Type defaultValueType = dv.GetType();

            Type thisType = GetType();

            for (Byte featureIndex = 1; featureIndex <= Plugin.FeatureCount; featureIndex++)
            {
                FieldInfo checkBoxField = thisType.GetField($"{Constants.Feature}{featureIndex}CheckBox", BindingFlags.NonPublic | BindingFlags.Instance);

                CheckBox checkBoxFieldValue = (CheckBox)(checkBoxField.GetValue(this));

                FieldInfo labelField = defaultValueType.GetField($"{Constants.Feature}{featureIndex}{Constants.LabelSuffix}", BindingFlags.Public | BindingFlags.Instance);

                String labelFieldValue = (String)(labelField.GetValue(dv));

                checkBoxFieldValue.Text = labelFieldValue;

                FieldInfo valueField = defaultValueType.GetField($"{Constants.Feature}{featureIndex}", BindingFlags.Public | BindingFlags.Instance);

                checkBoxFieldValue.Checked = (Boolean)(valueField.GetValue(dv));

                FieldInfo textBoxField = thisType.GetField($"{Constants.Feature}{featureIndex}TextBox", BindingFlags.NonPublic | BindingFlags.Instance);

                TextBox textBoxFieldValue = (TextBox)(textBoxField.GetValue(this));

                textBoxFieldValue.Text = labelFieldValue;
            }
        }

        private void SetInvelosLabels()
        {
            SceneAccessCheckBox.Text = Texts.SceneAccess;
            PlayAllCheckBox.Text = Texts.PlayAll;
            FeatureTrailersCheckBox.Text = Texts.FeatureTrailers;
            BonusTrailerCheckBox.Text = Texts.BonusTrailer;
            FeaturettesCheckBox.Text = Texts.Featurettes;
            CommentaryCheckBox.Text = Texts.Commentary;
            DeletedScenesCheckBox.Text = Texts.DeletedScenes;
            InterviewsCheckBox.Text = Texts.Interviews;
            OuttakesBloopersCheckBox.Text = Texts.OuttakesBloopers;
            StoryboardComparisonsCheckBox.Text = Texts.StoryboardComparisons;
            GalleryCheckBox.Text = Texts.Gallery;
            ProductionNotesBiosCheckBox.Text = Texts.ProductionNotesBios;
            DVDROMContentCheckBox.Text = Texts.DVDROMContent;
            InteractiveGamesCheckBox.Text = Texts.InteractiveGames;
            MultiAngleCheckBox.Text = Texts.MultiAngle;
            MusicVideosCheckBox.Text = Texts.MusicVideos;
            THXCertifiedCheckBox.Text = Texts.THXCertified;
            ClosedCaptionedCheckBox.Text = Texts.ClosedCaptioned;
            DigitalCopyCheckBox.Text = Texts.DigitalCopy;
            PictureInPictureCheckBox.Text = Texts.PictureInPicture;
            BDLiveCheckBox.Text = Texts.BDLive;
            DBoxCheckBox.Text = Texts.DBox;
            CineChatCheckBox.Text = Texts.CineChat;
            MovieIQCheckBox.Text = Texts.MovieIQ;
        }

        private void SetComboBoxes()
        {
            Dictionary<Int32, CultureInfo> uiLanguages = new Dictionary<Int32, CultureInfo>(3);

            AddLanguage(uiLanguages, "en");
            AddLanguage(uiLanguages, "de");
            AddLanguage(uiLanguages, "fr");

            UiLanguageComboBox.DataSource = new BindingSource(uiLanguages, null);
            UiLanguageComboBox.DisplayMember = "Value";
            UiLanguageComboBox.ValueMember = "Key";
            UiLanguageComboBox.Text = Plugin.Settings.DefaultValues.UiLanguage.DisplayName;
        }

        private static void AddLanguage(Dictionary<Int32, CultureInfo> uiLanguages
            , String language)
        {
            CultureInfo ci = CultureInfo.GetCultureInfo(language);

            uiLanguages.Add(ci.LCID, ci);
        }

        private void SetLabels()
        {
            SetInvelosLabels();

            SetPluginLabels();

            #region Misc

            Text = Texts.Options;

            #region TabPages

            LabelTabPage.Text = Texts.Labels;
            ExcelColumnsTabPage.Text = Texts.ExcelColumns;
            MiscTabPage.Text = Texts.Misc;
            ExcelColumnsTabPage.Text = Texts.ExcelColumns;
            InvelosBasicsTabPage.Text = Texts.Basics;
            InvelosFeaturesTabPage.Text = Texts.Features;
            InvelosTabPage.Text = Texts.InvelosData;
            PluginTabPage.Text = Texts.PluginData;

            #endregion

            #region Labels

            UiLanguageLabel.Text = Texts.UiLanguage;
            ExportToCollectionXmlLabel.Text = Texts.ExportToCollectionXml;
            FilterCountLabel.Text = Texts.FilterCount;

            IdCheckBox.Text = Texts.Id;
            TitleCheckBox.Text = Texts.Title;
            EditionCheckBox.Text = Texts.Edition;
            SortTitleCheckBox.Text = Texts.SortTitle;

            #endregion

            #region Buttons

            SaveButton.Text = Texts.Save;
            DiscardButton.Text = Texts.Cancel;

            #endregion

            #endregion
        }

        private void SetPluginLabels()
        {
            Type thisType = GetType();

            for (Byte featureIndex = 1; featureIndex <= Plugin.FeatureCount; featureIndex++)
            {
                FieldInfo labelField = thisType.GetField($"{Constants.Feature}{featureIndex}{Constants.LabelSuffix}", BindingFlags.NonPublic | BindingFlags.Instance);

                Label labelFieldValue = (Label)(labelField.GetValue(this));

                labelFieldValue.Text = $"{Texts.Feature} {featureIndex}";
            }
        }

        private void OnDiscardButtonClick(Object sender
            , EventArgs e)
        {
            DialogResult = DialogResult.Cancel;

            Close();
        }

        private void OnSaveButtonClick(Object sender
            , EventArgs e)
        {
            SaveInvelosSettings();

            SavePluginSetting();

            DefaultValues dv = Plugin.Settings.DefaultValues;

            dv.FilterCount = (Byte)(FilterCountUpDown.Value);

            dv.ExportToCollectionXml = ExportToCollectionXmlCheckBox.Checked;

            CultureInfo uiLanguage = GetUiLanguage();

            dv.UiLanguage = uiLanguage;

            Texts.Culture = uiLanguage;

            MessageBoxTexts.Culture = uiLanguage;

            DialogResult = DialogResult.OK;

            Close();
        }

        private void SaveInvelosSettings()
        {
            DefaultValues dv = Plugin.Settings.DefaultValues;

            dv.Id = IdCheckBox.Checked;
            dv.Title = TitleCheckBox.Checked;
            dv.Edition = EditionCheckBox.Checked;
            dv.SortTitle = SortTitleCheckBox.Checked;

            dv.SceneAccess = SceneAccessCheckBox.Checked;
            dv.PlayAll = PlayAllCheckBox.Checked;
            dv.FeatureTrailers = FeatureTrailersCheckBox.Checked;
            dv.BonusTrailers = BonusTrailerCheckBox.Checked;

            dv.Featurettes = FeaturettesCheckBox.Checked;
            dv.Commentary = CommentaryCheckBox.Checked;
            dv.DeletedScenes = DeletedScenesCheckBox.Checked;
            dv.Interviews = InterviewsCheckBox.Checked;
            dv.OuttakesBloopers = OuttakesBloopersCheckBox.Checked;

            dv.StoryboardComparisons = StoryboardComparisonsCheckBox.Checked;
            dv.Gallery = GalleryCheckBox.Checked;
            dv.ProductionNotesBios = ProductionNotesBiosCheckBox.Checked;

            dv.DVDROMContent = DVDROMContentCheckBox.Checked;
            dv.InteractiveGames = InteractiveGamesCheckBox.Checked;
            dv.MultiAngle = MultiAngleCheckBox.Checked;
            dv.MusicVideos = MusicVideosCheckBox.Checked;

            dv.THXCertified = THXCertifiedCheckBox.Checked;
            dv.ClosedCaptioned = ClosedCaptionedCheckBox.Checked;

            dv.DigitalCopy = DigitalCopyCheckBox.Checked;
            dv.PictureInPicture = PictureInPictureCheckBox.Checked;
            dv.BDLive = BDLiveCheckBox.Checked;
            dv.DBox = DBoxCheckBox.Checked;
            dv.CineChat = CineChatCheckBox.Checked;
            dv.MovieIQ = MovieIQCheckBox.Checked;
        }

        private void SavePluginSetting()
        {
            DefaultValues dv = Plugin.Settings.DefaultValues;

            Type defaultValueType = dv.GetType();

            Type thisType = GetType();

            for (Byte featureIndex = 1; featureIndex <= Plugin.FeatureCount; featureIndex++)
            {
                FieldInfo checkBoxField = thisType.GetField($"{Constants.Feature}{featureIndex}CheckBox", BindingFlags.NonPublic | BindingFlags.Instance);

                CheckBox checkBoxFieldValue = (CheckBox)(checkBoxField.GetValue(this));

                FieldInfo valueField = defaultValueType.GetField($"{Constants.Feature}{featureIndex}", BindingFlags.Public | BindingFlags.Instance);

                valueField.SetValue(dv, checkBoxFieldValue.Checked);

                FieldInfo textBoxField = thisType.GetField($"{Constants.Feature}{featureIndex}TextBox", BindingFlags.NonPublic | BindingFlags.Instance);

                FieldInfo labelField = defaultValueType.GetField($"{Constants.Feature}{featureIndex}{Constants.LabelSuffix}", BindingFlags.Public | BindingFlags.Instance);

                TextBox textBoxFieldValue = (TextBox)(textBoxField.GetValue(this));

                labelField.SetValue(dv, textBoxFieldValue.Text);
            }
        }

        private CultureInfo GetUiLanguage()
        {
            KeyValuePair<Int32, CultureInfo> kvp = (KeyValuePair<Int32, CultureInfo>)(UiLanguageComboBox.SelectedItem);

            CultureInfo ci = kvp.Value;

            return (ci);
        }
    }
}