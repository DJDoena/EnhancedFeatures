using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;
using DoenaSoft.DVDProfiler.EnhancedFeatures.Resources;
using DoenaSoft.DVDProfiler.DVDProfilerHelper;
using Invelos.DVDProfilerPlugin;
using Microsoft.WindowsAPICodePack.Taskbar;
using System.Reflection;

namespace DoenaSoft.DVDProfiler.EnhancedFeatures
{
    internal sealed class CsvManager
    {
        private readonly Plugin Plugin;

        public CsvManager(Plugin plugin)
        {
            Plugin = plugin;
        }

        #region Export

        internal void Export(Boolean exportAll)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.AddExtension = true;
                sfd.DefaultExt = ".csv";
                sfd.Filter = "CSV (comma-separated values) files|*.csv";
                sfd.OverwritePrompt = true;
                sfd.RestoreDirectory = true;
                sfd.Title = Texts.SaveXmlFile;
                sfd.FileName = "EnhancedFeatures.csv";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;

                    using (ProgressWindow progressWindow = new ProgressWindow())
                    {
                        #region Progress

                        progressWindow.ProgressBar.Minimum = 0;
                        progressWindow.ProgressBar.Step = 1;
                        progressWindow.CanClose = false;

                        #endregion

                        CultureInfo currentCulture = Application.CurrentCulture;

                        String listSeparator = currentCulture.TextInfo.ListSeparator;

                        String dateFormat = currentCulture.DateTimeFormat.ShortDatePattern;

                        Object[] ids = (exportAll) ? ((Object[])(Plugin.Api.GetAllProfileIDs())) : ((Object[])(Plugin.Api.GetFlaggedProfileIDs()));

                        #region Progress

                        progressWindow.ProgressBar.Maximum = ids.Length;
                        progressWindow.Show();

                        if (TaskbarManager.IsPlatformSupported)
                        {
                            TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.Normal);
                            TaskbarManager.Instance.SetProgressValue(0, progressWindow.ProgressBar.Maximum);
                        }

                        Int32 onePercent = progressWindow.ProgressBar.Maximum / 100;

                        if ((progressWindow.ProgressBar.Maximum % 100) != 0)
                        {
                            onePercent++;
                        }

                        #endregion

                        try
                        {
                            using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write, FileShare.Read))
                            {
                                using (StreamWriter sw = new StreamWriter(fs, Encoding.UTF8))
                                {
                                    DefaultValues dv = Plugin.Settings.DefaultValues;

                                    #region Header

                                    WriteInvelosBasicsHeader(sw, dv, listSeparator);

                                    WriteInvelosFeaturesHeader(sw, dv, listSeparator);

                                    Type defaultValueType = dv.GetType();

                                    WritePluginHeader(sw, dv, defaultValueType, listSeparator);

                                    sw.WriteLine();

                                    #endregion

                                    for (Int32 profileIndex = 0; profileIndex < ids.Length; profileIndex++)
                                    {
                                        #region Row

                                        String id = ids[profileIndex].ToString();

                                        IDVDInfo profile;
                                        Plugin.Api.DVDByProfileID(out profile, id, PluginConstants.DATASEC_AllSections, 0);

                                        WriteInvelosBasicsRow(sw, profile, dv, listSeparator);

                                        WriteInvelosFeaturesRow(sw, profile, dv, listSeparator);

                                        WritePluginRow(sw, profile, dv, defaultValueType, listSeparator);

                                        sw.WriteLine();

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

                                        #endregion
                                    }
                                }
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

        private static void WriteInvelosBasicsHeader(StreamWriter sw
            , DefaultValues dv
            , String listSeparator)
        {
            if (dv.Id)
            {
                WriteText(sw, Texts.Id, listSeparator);
            }

            if (dv.Title)
            {
                WriteText(sw, Texts.Title, listSeparator);
            }

            if (dv.Edition)
            {
                WriteText(sw, Texts.Edition, listSeparator);
            }

            if (dv.SortTitle)
            {
                WriteText(sw, Texts.SortTitle, listSeparator);
            }
        }

        private void WriteInvelosFeaturesHeader(StreamWriter sw
            , DefaultValues dv
            , String listSeparator)
        {
            if (dv.SceneAccess)
            {
                WriteText(sw, Texts.SceneAccess, listSeparator);
            }

            if (dv.PlayAll)
            {
                WriteText(sw, Texts.PlayAll, listSeparator);
            }

            if (dv.FeatureTrailers)
            {
                WriteText(sw, Texts.FeatureTrailers, listSeparator);
            }

            if (dv.BonusTrailers)
            {
                WriteText(sw, Texts.BonusTrailer, listSeparator);
            }

            if (dv.Featurettes)
            {
                WriteText(sw, Texts.Featurettes, listSeparator);
            }

            if (dv.Commentary)
            {
                WriteText(sw, Texts.Commentary, listSeparator);
            }

            if (dv.DeletedScenes)
            {
                WriteText(sw, Texts.DeletedScenes, listSeparator);
            }

            if (dv.Interviews)
            {
                WriteText(sw, Texts.Interviews, listSeparator);
            }

            if (dv.OuttakesBloopers)
            {
                WriteText(sw, Texts.OuttakesBloopers, listSeparator);
            }

            if (dv.StoryboardComparisons)
            {
                WriteText(sw, Texts.StoryboardComparisons, listSeparator);
            }

            if (dv.Gallery)
            {
                WriteText(sw, Texts.Gallery, listSeparator);
            }

            if (dv.ProductionNotesBios)
            {
                WriteText(sw, Texts.ProductionNotesBios, listSeparator);
            }

            if (dv.DVDROMContent)
            {
                WriteText(sw, Texts.DVDROMContent, listSeparator);
            }

            if (dv.InteractiveGames)
            {
                WriteText(sw, Texts.InteractiveGames, listSeparator);
            }

            if (dv.MultiAngle)
            {
                WriteText(sw, Texts.MultiAngle, listSeparator);
            }

            if (dv.MusicVideos)
            {
                WriteText(sw, Texts.MusicVideos, listSeparator);
            }

            if (dv.THXCertified)
            {
                WriteText(sw, Texts.THXCertified, listSeparator);
            }

            if (dv.ClosedCaptioned)
            {
                WriteText(sw, Texts.ClosedCaptioned, listSeparator);
            }

            if (dv.DigitalCopy)
            {
                WriteText(sw, Texts.DigitalCopy, listSeparator);
            }

            if (dv.PictureInPicture)
            {
                WriteText(sw, Texts.PictureInPicture, listSeparator);
            }

            if (dv.BDLive)
            {
                WriteText(sw, Texts.BDLive, listSeparator);
            }

            if (dv.DBox)
            {
                WriteText(sw, Texts.DBox, listSeparator);
            }

            if (dv.CineChat)
            {
                WriteText(sw, Texts.CineChat, listSeparator);
            }

            if (dv.MovieIQ)
            {
                WriteText(sw, Texts.MovieIQ, listSeparator);
            }
        }

        private static void WritePluginHeader(StreamWriter sw
            , DefaultValues dv
            , Type defaultValueType
            , String listSeparator)
        {
            for (Byte featureIndex = 1; featureIndex <= Plugin.FeatureCount; featureIndex++)
            {
                FieldInfo valueField = defaultValueType.GetField($"{Constants.Feature}{featureIndex}", BindingFlags.Public | BindingFlags.Instance);

                Boolean valueFieldValue = (Boolean)(valueField.GetValue(dv));

                if (valueFieldValue)
                {
                    FieldInfo labelField = defaultValueType.GetField($"{Constants.Feature}{featureIndex}{Constants.LabelSuffix}", BindingFlags.Public | BindingFlags.Instance);

                    WriteText(sw, (String)(labelField.GetValue(dv)), listSeparator);
                }
            }
        }

        private static void WriteInvelosBasicsRow(StreamWriter sw
            , IDVDInfo profile
            , DefaultValues dv
            , String listSeparator)
        {
            if (dv.Id)
            {
                WriteText(sw, profile.GetFormattedProfileID(), listSeparator);
            }

            if (dv.Title)
            {
                WriteText(sw, profile.GetTitle(), listSeparator);
            }

            if (dv.Edition)
            {
                WriteText(sw, profile.GetEdition(), listSeparator);
            }

            if (dv.SortTitle)
            {
                WriteText(sw, profile.GetSortTitle(), listSeparator);
            }
        }

        private void WriteInvelosFeaturesRow(StreamWriter sw
            , IDVDInfo profile
            , DefaultValues dv
            , String listSeparator)
        {
            if (dv.SceneAccess)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_SceneAccess), listSeparator);
            }

            if (dv.PlayAll)
            {
                //WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_PlayAll), listSeparator);
                WriteBoolean(sw, false, listSeparator);
            }

            if (dv.FeatureTrailers)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_Trailer), listSeparator);
            }

            if (dv.BonusTrailers)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_BonusTrailers), listSeparator);
            }

            if (dv.Featurettes)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_Documentary), listSeparator);
            }

            if (dv.Commentary)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_Commentary), listSeparator);
            }

            if (dv.DeletedScenes)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_DeletedScenes), listSeparator);
            }

            if (dv.Interviews)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_Interviews), listSeparator);
            }

            if (dv.OuttakesBloopers)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_Bloopers), listSeparator);
            }

            if (dv.StoryboardComparisons)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_StoryboardComps), listSeparator);
            }

            if (dv.Gallery)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_Gallery), listSeparator);
            }

            if (dv.ProductionNotesBios)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_ProductionNotes), listSeparator);
            }

            if (dv.DVDROMContent)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_DVDROMContent), listSeparator);
            }

            if (dv.InteractiveGames)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_InteractiveGame), listSeparator);
            }

            if (dv.MultiAngle)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_MultiAngle), listSeparator);
            }

            if (dv.MusicVideos)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_MusicVideos), listSeparator);
            }

            if (dv.THXCertified)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_THX), listSeparator);
            }

            if (dv.ClosedCaptioned)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_ClosedCaptioned), listSeparator);
            }

            if (dv.DigitalCopy)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_DigitalCopy), listSeparator);
            }

            if (dv.PictureInPicture)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_PIP), listSeparator);
            }

            if (dv.BDLive)
            {
                WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_BDLive), listSeparator);
            }

            if (dv.DBox)
            {
                //WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_DBox), listSeparator);
                WriteBoolean(sw, false, listSeparator);
            }

            if (dv.CineChat)
            {
                //WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_CineChat), listSeparator);
                WriteBoolean(sw, false, listSeparator);
            }

            if (dv.MovieIQ)
            {
                //WriteBoolean(sw, profile.GetFeatureByID(PluginConstants.FEATURE_MovieIQ), listSeparator);
                WriteBoolean(sw, false, listSeparator);
            }
        }

        private static void WritePluginRow(StreamWriter sw
            , IDVDInfo profile
            , DefaultValues dv
            , Type defaultValueType
            , String listSeparator)
        {
            FeatureManager featureManager = new FeatureManager(profile);

            for (Byte featureIndex = 1; featureIndex <= Plugin.FeatureCount; featureIndex++)
            {
                FieldInfo valueField = defaultValueType.GetField($"{Constants.Feature}{featureIndex}", BindingFlags.Public | BindingFlags.Instance);

                Boolean valueFieldValue = (Boolean)(valueField.GetValue(dv));

                if (valueFieldValue)
                {
                    Boolean value = featureManager.GetFeature(featureIndex);

                    WriteBoolean(sw, value, listSeparator);
                }
            }
        }

        private static void WriteBoolean(StreamWriter sw
            , Boolean value
            , String listSeparator)
        {
            String text = value ? "X" : String.Empty;

            WriteText(sw, text, listSeparator);
        }

        private static void WriteText(StreamWriter sw
            , String text
            , String listSeparator)
        {
            sw.Write("\"");

            if (String.IsNullOrEmpty(text) == false)
            {
                text = text.Replace("\"", "\"\"");
            }

            sw.Write(text);
            sw.Write("\"");
            sw.Write(listSeparator);
        }

        #endregion
    }
}