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

        private Label[] FeatureLabels;

        private TextBox[] FeatureTextBoxes;

        private CheckBox[] FeatureCheckBoxes;

        public SettingsForm(Plugin plugin)
        {
            Plugin = plugin;

            InitializeComponent();

            SetControls();

            SetSettings();

            SetLabels();

            SetComboBoxes();
        }

        #region SetControls

        private void SetControls()
        {
            SetLabelTabPageControls();

            SetPluginTabPageControls();
        }

        private void SetLabelTabPageControls()
        {
            const Byte FeatureCount = Plugin.FeatureCount;

            FeatureLabels = new Label[FeatureCount];

            FeatureTextBoxes = new TextBox[FeatureCount];

            for (Byte featureIndex = 1; featureIndex <= FeatureCount; featureIndex++)
            {
                const Int32 Offset = 26;

                const Byte Half = FeatureCount / 2;

                Int32 left;
                Int32 top;
                if (featureIndex <= Half)
                {
                    left = 6;

                    top = 6 + ((featureIndex - 1) * Offset);
                }
                else
                {
                    left = 336;

                    top = 6 + ((featureIndex - 1 - Half) * Offset);
                }

                AddFeatureLabel(featureIndex, left, top + 3);

                AddFeatureTextBoxes(featureIndex, left + 70, top);
            }

            LabelTabPage.Controls.AddRange(FeatureLabels);
            LabelTabPage.Controls.AddRange(FeatureTextBoxes);
        }

        private void AddFeatureLabel(Byte featureIndex
            , Int32 left
            , Int32 top)
        {
            Label label = new Label();

            label.AutoSize = true;
            label.Location = new System.Drawing.Point(left, top);
            label.Name = $"Feature{featureIndex}{Constants.LabelSuffix}";
            label.Size = new System.Drawing.Size(55, 13);
            label.TabIndex = (featureIndex - 1) * 2;

            FeatureLabels[featureIndex - 1] = label;
        }

        private void AddFeatureTextBoxes(Byte featureIndex
            , Int32 left
            , Int32 top)
        {
            TextBox textBox = new TextBox();

            textBox.Location = new System.Drawing.Point(left, top);
            textBox.MaxLength = 30;
            textBox.Name = $"Feature{featureIndex}TextBox";
            textBox.Size = new System.Drawing.Size(200, 20);
            textBox.TabIndex = (featureIndex - 1) * 2 + 1;

            FeatureTextBoxes[featureIndex - 1] = textBox;
        }

        private void SetPluginTabPageControls()
        {
            const Byte FeatureCount = Plugin.FeatureCount;

            FeatureCheckBoxes = new CheckBox[FeatureCount];

            for (Byte featureIndex = 1; featureIndex <= FeatureCount; featureIndex++)
            {
                const Int32 Offset = 23;

                const Byte Half = FeatureCount / 2;

                Int32 left;
                Int32 top;
                if (featureIndex <= Half)
                {
                    left = 6;

                    top = 6 + ((featureIndex - 1) * Offset);
                }
                else
                {
                    left = 306;

                    top = 6 + ((featureIndex - 1 - Half) * Offset);
                }

                AddFeatureCheckBox(featureIndex, left, top);
            }

            PluginTabPage.Controls.AddRange(FeatureCheckBoxes);
        }

        private void AddFeatureCheckBox(Byte featureIndex
            , Int32 left
            , Int32 top)
        {
            CheckBox checkBox = new CheckBox();

            checkBox.AutoSize = true;
            checkBox.Location = new System.Drawing.Point(left, top);
            checkBox.Name = $"Feature{featureIndex}CheckBox";
            checkBox.Size = new System.Drawing.Size(62, 17);
            checkBox.TabIndex = featureIndex - 1;
            checkBox.UseVisualStyleBackColor = true;

            FeatureCheckBoxes[featureIndex - 1] = checkBox;
        }

        #endregion

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

            for (Byte featureIndex = 1; featureIndex <= Plugin.FeatureCount; featureIndex++)
            {
                CheckBox checkBox = FeatureCheckBoxes[featureIndex - 1];

                String label = dv.FeatureLabels[featureIndex];

                checkBox.Text = label;

                checkBox.Checked = dv.ExcelFeatures[featureIndex];

                FeatureTextBoxes[featureIndex - 1].Text = label;
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
            for (Byte featureIndex = 1; featureIndex <= Plugin.FeatureCount; featureIndex++)
            {
                FeatureLabels[featureIndex - 1].Text = $"{Texts.Feature} {featureIndex}";
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

            SavePluginSettings();

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

        private void SavePluginSettings()
        {
            DefaultValues dv = Plugin.Settings.DefaultValues;

            for (Byte featureIndex = 1; featureIndex <= Plugin.FeatureCount; featureIndex++)
            {
                dv.ExcelFeatures[featureIndex] = FeatureCheckBoxes[featureIndex - 1].Checked;

                dv.FeatureLabels[featureIndex] = FeatureTextBoxes[featureIndex - 1].Text;
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