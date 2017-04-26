using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using DoenaSoft.DVDProfiler.DVDProfilerHelper;
using DoenaSoft.DVDProfiler.EnhancedFeatures.Resources;
using Invelos.DVDProfilerPlugin;

namespace DoenaSoft.DVDProfiler.EnhancedFeatures
{
    internal sealed partial class MainForm : Form
    {
        private readonly Plugin Plugin;

        private readonly IDVDInfo Profile;

        private readonly FeatureManager FeatureManager;

        private Boolean DataChanged;

        private CheckBox[] FeatureCheckBoxes;

        internal MainForm(Plugin plugin
            , IDVDInfo profile)
        {
            Plugin = plugin;
            Profile = profile;

            InitializeComponent();

            SetPluginTabPageControls();

            FeatureManager = new FeatureManager(Profile);

            SetData();

            SetReadOnlies();

            DataChanged = false;
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

        private void SetReadOnlies()
        {
            if (Plugin.IsRemoteAccess)
            {
                ImportFromXMLToolStripMenuItem.Enabled = false;

                PasteAllToolStripMenuItem.Enabled = false;

                SaveButton.Enabled = false;

                SetControlsReadonly(Controls);
            }
        }

        private void SetControlsReadonly(Control.ControlCollection controls)
        {
            if (controls != null)
            {
                foreach (Control control in controls)
                {
                    if (control is CheckBox)
                    {
                        ((CheckBox)control).Enabled = true;
                    }
                    else
                    {
                        SetControlsReadonly(control.Controls);
                    }
                }
            }
        }

        #region SetData

        private void SetData()
        {
            SetInvelosCheckeds();

            SetPluginCheckeds();

            SetLabels();
        }

        private void SetLabels()
        {
            SetInvelosLabels();

            SetPluginLabels();

            SetMiscLabels();
        }

        private void SetPluginLabels()
        {
            DefaultValues dv = Plugin.Settings.DefaultValues;

            for (Byte featureIndex = 1; featureIndex <= Plugin.FeatureCount; featureIndex++)
            {
                FeatureCheckBoxes[featureIndex - 1].Text = dv.FeatureLabels[featureIndex];
            }
        }

        private void SetPluginCheckeds()
        {
            DefaultValues dv = Plugin.Settings.DefaultValues;

            for (Byte featureIndex = 1; featureIndex <= Plugin.FeatureCount; featureIndex++)
            {
                FeatureCheckBoxes[featureIndex - 1].Checked = FeatureManager.GetFeature(featureIndex);
            }
        }

        #region SetInvelosData

        private void SetInvelosCheckeds()
        {
            SceneAccessCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_SceneAccess);
            //PlayAllCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_PlayAll);
            FeatureTrailersCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_Trailer);
            BonusTrailerCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_BonusTrailers);
            FeaturettesCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_Documentary);
            CommentaryCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_Commentary);
            DeletedScenesCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_DeletedScenes);
            InterviewsCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_Interviews);
            OuttakesBloopersCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_Bloopers);
            StoryboardComparisonsCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_StoryboardComps);
            GalleryCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_Gallery);
            ProductionNotesBiosCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_ProductionNotes);
            DVDROMContentCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_DVDROMContent);
            InteractiveGamesCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_InteractiveGame);
            MultiAngleCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_MultiAngle);
            MusicVideosCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_MusicVideos);
            THXCertifiedCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_THX);
            ClosedCaptionedCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_ClosedCaptioned);
            DigitalCopyCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_DigitalCopy);
            PictureInPictureCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_PIP);
            BDLiveCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_BDLive);
            //DBoxCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_DBox);
            //CineChatCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_CineChat);
            //MovieIQCheckBox.Checked = Profile.GetFeatureByID(PluginConstants.FEATURE_MovieIQ);
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

        #endregion

        private void SetMiscLabels()
        {
            #region Menu

            EditToolStripMenuItem.Text = Texts.Edit;
            CopyAllToolStripMenuItem.Text = Texts.CopyAllToClipboard;
            PasteAllToolStripMenuItem.Text = Texts.PasteAllFromClipboard;

            ToolsToolStripMenuItem.Text = Texts.Tools;
            OptionsToolStripMenuItem.Text = Texts.Options;
            ExportToXMLToolStripMenuItem.Text = Texts.ExportToXml;
            ImportFromXMLToolStripMenuItem.Text = Texts.ImportFromXml;
            ExportOptionsToolStripMenuItem.Text = Texts.ExportOptions;
            ImportOptionsToolStripMenuItem.Text = Texts.ImportOptions;

            HelpToolStripMenuItem.Text = Texts.Help;
            CheckForUpdatesToolStripMenuItem.Text = Texts.CheckForUpdates;
            AboutToolStripMenuItem.Text = Texts.About;

            #endregion

            #region TabPages

            InvelosTabPage.Text = Texts.InvelosData;
            PluginTabPage.Text = Texts.PluginData;

            #endregion

            #region Buttons

            SaveButton.Text = Texts.Save;
            DiscardButton.Text = Texts.Cancel;

            #endregion
        }

        #endregion

        private void OnSaveButtonClick(Object sender
            , EventArgs e)
        {
            DefaultValues dv = Plugin.Settings.DefaultValues;

            for (Byte featureIndex = 1; featureIndex <= Plugin.FeatureCount; featureIndex++)
            {
                FeatureManager.SetFeature(featureIndex, FeatureCheckBoxes[featureIndex - 1].Checked);
            }

            Plugin.Api.SaveDVDToCollection(Profile);

            Plugin.Api.ReloadCurrentDVD();

            DataChanged = false;

            Close();
        }

        private void OnDiscardButtonClick(Object sender, EventArgs e)
        {
            Close();
        }

        private void OnOptionsToolStripMenuItemClick(Object sender
            , EventArgs e)
        {
            using (SettingsForm form = new SettingsForm(Plugin))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    SetLabels();

                    Plugin.RegisterCustomFields();
                }
            }
        }

        private void OnExportToXMLToolStripMenuItemClick(Object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.AddExtension = true;
                sfd.DefaultExt = ".xml";
                sfd.Filter = "XML files|*.xml";
                sfd.OverwritePrompt = true;
                sfd.RestoreDirectory = true;
                sfd.Title = Texts.SaveXmlFile;
                sfd.FileName = "EnhancedFeatures." + Profile.GetProfileID() + ".xml";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    EnhancedFeatures ef = GetEnhancedFeaturesForXmlStructure();

                    try
                    {
                        Serializer<EnhancedFeatures>.Serialize(sfd.FileName, ef);
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

        private EnhancedFeatures GetEnhancedFeaturesForXmlStructure()
        {
            DefaultValues dv = Plugin.Settings.DefaultValues;

            EnhancedFeatures ef = new EnhancedFeatures();

            const Byte FeatureCount = Plugin.FeatureCount;

            List<Feature> features = new List<Feature>();

            for (Byte featureIndex = 1; featureIndex <= FeatureCount; featureIndex++)
            {
                Boolean isChecked = FeatureCheckBoxes[featureIndex - 1].Checked;

                if (isChecked)
                {
                    Feature feature = new Feature();

                    feature.Index = featureIndex;

                    feature.Value = isChecked;

                    feature.DisplayName = dv.FeatureLabels[featureIndex];

                    features.Add(feature);
                }
            }

            ef.Feature = features.ToArray();

            return (ef);
        }

        private void OnImportFromXMLToolStripMenuItemClick(Object sender
            , EventArgs e)
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
                    EnhancedFeatures ef = null;

                    try
                    {
                        ef = Serializer<EnhancedFeatures>.Deserialize(ofd.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(String.Format(MessageBoxTexts.FileCantBeRead, ofd.FileName, ex.Message)
                           , MessageBoxTexts.ErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    if (ef?.Feature != null)
                    {
                        SetEnhancedFeaturesFromXmlStructure(ef.Feature);
                    }
                }
            }
        }

        private void SetEnhancedFeaturesFromXmlStructure(Feature[] features)
        {
            foreach (Feature feature in features)
            {
                FeatureCheckBoxes[feature.Index - 1].Checked = feature.Value;
            }
        }

        private void OnCheckForUpdatesToolStripMenuItemClick(Object sender, EventArgs e)
        {
            OnlineAccess.Init("Doena Soft.", "EnhancedFeatures");
            OnlineAccess.CheckForNewVersion("http://doena-soft.de/dvdprofiler/3.9.5/versions.xml", this, "EnhancedFeatures", GetType().Assembly);
        }

        private void OnAboutToolStripMenuItemClick(Object sender, EventArgs e)
        {
            using (AboutBox aboutBox = new AboutBox(GetType().Assembly))
            {
                aboutBox.ShowDialog();
            }
        }

        private void OnDataChanged(Object sender
            , EventArgs e)
        {
            DataChanged = true;
        }

        private void OnFormClosing(Object sender
            , FormClosingEventArgs e)
        {
            if (DataChanged)
            {
                if (MessageBox.Show(MessageBoxTexts.AbandonChangesText, MessageBoxTexts.AbandonChangesHeader, MessageBoxButtons.YesNo
                    , MessageBoxIcon.Question) == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
        }

        private void OnImportOptionsToolStripMenuItemClick(Object sender
            , EventArgs e)
        {
            Plugin.ImportOptions();

            SetPluginLabels();
        }

        private void OnExportOptionsToolStripMenuItemClick(Object sender
            , EventArgs e)
        {
            Plugin.ExportOptions();
        }

        private void OnCopyAllToolStripMenuItemClick(Object sender
            , EventArgs e)
        {
            EnhancedFeatures ef = GetEnhancedFeaturesForXmlStructure();

            String xml = Serializer<EnhancedFeatures>.ToString(ef);

            try
            {
                Clipboard.SetDataObject(xml, true, 4, 250);
            }
            catch
            {
                MessageBox.Show(MessageBoxTexts.CopyToClipboardFailed, MessageBoxTexts.ErrorHeader
                 , MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void OnPasteAllToolStripMenuItemClick(Object sender
            , EventArgs e)
        {
            EnhancedFeatures ef = null;

            try
            {
                String xml = Clipboard.GetText();

                ef = Serializer<EnhancedFeatures>.FromString(xml);
            }
            catch
            {
                MessageBox.Show(MessageBoxTexts.PasteFromClipboardFailed, MessageBoxTexts.ErrorHeader
                    , MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            if (ef?.Feature != null)
            {
                SetEnhancedFeaturesFromXmlStructure(ef.Feature);
            }
        }
    }
}