using System;
using System.Collections.Generic;
using System.Windows.Forms;
using DoenaSoft.DVDProfiler.DVDProfilerHelper;
using DoenaSoft.DVDProfiler.EnhancedFeatures.Resources;
using DoenaSoft.ToolBox.Generics;
using Invelos.DVDProfilerPlugin;
using Microsoft.WindowsAPICodePack.Taskbar;

namespace DoenaSoft.DVDProfiler.EnhancedFeatures
{
    internal sealed class XmlManager
    {
        private readonly Plugin Plugin;

        public XmlManager(Plugin plugin)
        {
            Plugin = plugin;
        }

        #region Export

        internal void Export(Boolean exportAll)
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.AddExtension = true;
                sfd.DefaultExt = ".xml";
                sfd.Filter = "XML files|*.xml";
                sfd.OverwritePrompt = true;
                sfd.RestoreDirectory = true;
                sfd.Title = Texts.SaveXmlFile;
                sfd.FileName = "EnhancedFeatures.xml";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Object[] ids;
                    EnhancedFeaturesList efs;

                    Cursor.Current = Cursors.WaitCursor;
                    using (var progressWindow = new ProgressWindow())
                    {
                        #region Progress

                        Int32 onePercent;

                        progressWindow.ProgressBar.Minimum = 0;
                        progressWindow.ProgressBar.Step = 1;
                        progressWindow.CanClose = false;

                        #endregion

                        ids = this.GetProfileIds(exportAll);

                        efs = new EnhancedFeaturesList();
                        efs.Profiles = new Profile[ids.Length];

                        #region Progress

                        progressWindow.ProgressBar.Maximum = ids.Length;
                        progressWindow.Show();
                        if (TaskbarManager.IsPlatformSupported)
                        {
                            TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.Normal);
                            TaskbarManager.Instance.SetProgressValue(0, progressWindow.ProgressBar.Maximum);
                        }
                        onePercent = progressWindow.ProgressBar.Maximum / 100;
                        if ((progressWindow.ProgressBar.Maximum % 100) != 0)
                        {
                            onePercent++;
                        }

                        #endregion

                        for (var i = 0; i < ids.Length; i++)
                        {
                            String id;
                            IDVDInfo profile;

                            id = ids[i].ToString();
                            Plugin.Api.DVDByProfileID(out profile, id, PluginConstants.DATASEC_AllSections, 0);
                            efs.Profiles[i] = this.GetXmlProfile(profile);

                            #region Progress

                            progressWindow.ProgressBar.PerformStep();
                            if (TaskbarManager.IsPlatformSupported)
                            {
                                TaskbarManager.Instance.SetProgressValue(progressWindow.ProgressBar.Value, progressWindow.ProgressBar.Maximum);
                            }
                            if ((progressWindow.ProgressBar.Value % onePercent) == 0)
                            {
                                Application.DoEvents();
                            }

                            #endregion
                        }

                        try
                        {
                            Serializer<EnhancedFeaturesList>.Serialize(sfd.FileName, efs);

                            #region Progress

                            Application.DoEvents();
                            if (TaskbarManager.IsPlatformSupported)
                            {
                                TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.NoProgress);
                            }
                            progressWindow.CanClose = true;
                            progressWindow.Close();

                            #endregion

                            MessageBox.Show(String.Format(MessageBoxTexts.DoneWithNumber, ids.Length, MessageBoxTexts.Exported)
                                , MessageBoxTexts.InformationHeader, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(String.Format(MessageBoxTexts.FileCantBeWritten, sfd.FileName, ex.Message)
                                , MessageBoxTexts.ErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            #region Progress

                            if (progressWindow.Visible)
                            {
                                if (TaskbarManager.IsPlatformSupported)
                                {
                                    TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.NoProgress);
                                }
                                progressWindow.CanClose = true;
                                progressWindow.Close();
                            }

                            #endregion
                        }
                    }

                    Cursor.Current = Cursors.Default;
                }
            }
        }

        private Profile GetXmlProfile(IDVDInfo profile)
        {
            var dv = Plugin.Settings.DefaultValues;

            var featureManager = new FeatureManager(profile);

            var xmlProfile = new Profile();

            xmlProfile.Id = profile.GetProfileID();
            xmlProfile.Title = profile.GetTitle();

            const Byte FeatureCount = Plugin.FeatureCount;

            var features = new List<Feature>(FeatureCount);

            for (Byte featureIndex = 1; featureIndex <= FeatureCount; featureIndex++)
            {
                var value = featureManager.GetFeature(featureIndex);

                if (value)
                {
                    var feature = new Feature();

                    feature.Index = featureIndex;

                    feature.Value = value;

                    feature.DisplayName = dv.FeatureLabels[featureIndex];

                    features.Add(feature);
                }
            }

            xmlProfile.EnhancedFeatures = features.ToArray();

            return (xmlProfile);
        }

        #endregion

        #region Import

        internal void Import(Boolean importAll)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.CheckFileExists = true;
                ofd.Filter = "XML files|*.xml";
                ofd.Multiselect = false;
                ofd.RestoreDirectory = true;
                ofd.Title = Texts.LoadXmlFile;

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;

                    EnhancedFeaturesList efs = null;

                    try
                    {
                        efs = Serializer<EnhancedFeaturesList>.Deserialize(ofd.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(String.Format(MessageBoxTexts.FileCantBeRead, ofd.FileName, ex.Message)
                           , MessageBoxTexts.ErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    if (efs != null)
                    {
                        var count = 0;

                        if (efs.Profiles?.Length > 0)
                        {
                            using (var progressWindow = new ProgressWindow())
                            {
                                #region Progress

                                progressWindow.ProgressBar.Minimum = 0;
                                progressWindow.ProgressBar.Step = 1;
                                progressWindow.CanClose = false;

                                #endregion

                                var ids = this.GetProfileIds(importAll);

                                var profileIds = new Dictionary<String, Boolean>(ids.Length);

                                #region Progress

                                progressWindow.ProgressBar.Maximum = ids.Length;
                                progressWindow.Show();

                                if (TaskbarManager.IsPlatformSupported)
                                {
                                    TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.Normal);
                                    TaskbarManager.Instance.SetProgressValue(0, progressWindow.ProgressBar.Maximum);
                                }

                                var onePercent = progressWindow.ProgressBar.Maximum / 100;

                                if ((progressWindow.ProgressBar.Maximum % 100) != 0)
                                {
                                    onePercent++;
                                }

                                #endregion

                                for (var i = 0; i < ids.Length; i++)
                                {
                                    profileIds.Add(ids[i].ToString(), true);
                                }

                                foreach (var xmlProfile in efs.Profiles)
                                {
                                    if ((xmlProfile?.EnhancedFeatures != null) && (profileIds.ContainsKey(xmlProfile.Id)))
                                    {
                                        var profile = Plugin.Api.CreateDVD();

                                        profile.SetProfileID(xmlProfile.Id);

                                        var featureManager = new FeatureManager(profile);

                                        var features = xmlProfile.EnhancedFeatures;

                                        foreach (var feature in features)
                                        {
                                            featureManager.SetFeature(feature.Index, feature.Value);
                                        }

                                        count++;
                                    }

                                    #region Progress

                                    progressWindow.ProgressBar.PerformStep();

                                    if (TaskbarManager.IsPlatformSupported)
                                    {
                                        TaskbarManager.Instance.SetProgressValue(progressWindow.ProgressBar.Value, progressWindow.ProgressBar.Maximum);
                                    }

                                    if ((progressWindow.ProgressBar.Value % onePercent) == 0)
                                    {
                                        Application.DoEvents();
                                    }

                                    #endregion
                                }

                                #region Progress

                                Application.DoEvents();

                                if (TaskbarManager.IsPlatformSupported)
                                {
                                    TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.NoProgress);
                                }

                                progressWindow.CanClose = true;
                                progressWindow.Close();

                                #endregion
                            }
                        }

                        MessageBox.Show(String.Format(MessageBoxTexts.DoneWithNumber, count, MessageBoxTexts.Imported)
                                , MessageBoxTexts.InformationHeader, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    Cursor.Current = Cursors.Default;
                }
            }
        }

        #endregion

        private Object[] GetProfileIds(Boolean allIds)
            => ((allIds) ? ((Object[])(Plugin.Api.GetAllProfileIDs())) : ((Object[])(Plugin.Api.GetFlaggedProfileIDs())));
    }
}