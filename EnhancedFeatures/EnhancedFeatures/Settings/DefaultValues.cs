using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Serialization;
using DoenaSoft.DVDProfiler.EnhancedFeatures.Resources;

namespace DoenaSoft.DVDProfiler.EnhancedFeatures
{
    [ComVisible(false)]
    public sealed class DefaultValues
    {
        #region Labels

        public String Feature1Label = Texts.Feature1;

        public String Feature2Label = Texts.Feature2;

        public String Feature3Label = Texts.Feature3;

        public String Feature4Label = Texts.Feature4;

        public String Feature5Label = Texts.Feature5;

        public String Feature6Label = Texts.Feature6;

        public String Feature7Label = Texts.Feature7;

        public String Feature8Label = Texts.Feature8;

        public String Feature9Label = Texts.Feature9;

        public String Feature10Label = Texts.Feature10;

        public String Feature11Label = Texts.Feature11;

        public String Feature12Label = Texts.Feature12;

        public String Feature13Label = Texts.Feature13;

        public String Feature14Label = Texts.Feature14;

        public String Feature15Label = Texts.Feature15;

        public String Feature16Label = Texts.Feature16;

        public String Feature17Label = Texts.Feature17;

        public String Feature18Label = Texts.Feature18;

        public String Feature19Label = Texts.Feature19;

        public String Feature20Label = Texts.Feature20;

        public String Feature21Label = Texts.Feature21;

        public String Feature22Label = Texts.Feature22;

        public String Feature23Label = Texts.Feature23;

        public String Feature24Label = Texts.Feature24;

        public String Feature25Label = Texts.Feature25;

        public String Feature26Label = Texts.Feature26;

        public String Feature27Label = Texts.Feature27;

        public String Feature28Label = Texts.Feature28;

        public String Feature29Label = Texts.Feature29;

        public String Feature30Label = Texts.Feature30;

        #endregion

        #region InvelosData

        public Boolean Id = true;

        public Boolean Title = true;

        public Boolean SortTitle = true;

        public Boolean Edition = true;

        public Boolean SceneAccess = false;

        public Boolean PlayAll = false;

        public Boolean FeatureTrailers = false;

        public Boolean BonusTrailers = false;

        public Boolean Featurettes = false;

        public Boolean Commentary = false;

        public Boolean DeletedScenes = false;

        public Boolean ProductionNotesBios = false;

        public Boolean Interviews = false;

        public Boolean OuttakesBloopers = false;

        public Boolean StoryboardComparisons = false;

        public Boolean Gallery = false;

        public Boolean DVDROMContent = false;

        public Boolean InteractiveGames = false;

        public Boolean MultiAngle = false;

        public Boolean MusicVideos = false;

        public Boolean THXCertified = false;

        public Boolean ClosedCaptioned = false;

        public Boolean DigitalCopy = false;

        public Boolean PictureInPicture = false;

        public Boolean BDLive = false;

        public Boolean DBox = false;

        public Boolean CineChat = false;

        public Boolean MovieIQ = false;

        #endregion

        #region  PluginData

        public Boolean Feature1 = false;

        public Boolean Feature2 = false;

        public Boolean Feature3 = false;

        public Boolean Feature4 = false;

        public Boolean Feature5 = false;

        public Boolean Feature6 = false;

        public Boolean Feature7 = false;

        public Boolean Feature8 = false;

        public Boolean Feature9 = false;

        public Boolean Feature10 = false;

        public Boolean Feature11 = false;

        public Boolean Feature12 = false;

        public Boolean Feature13 = false;

        public Boolean Feature14 = false;

        public Boolean Feature15 = false;

        public Boolean Feature16 = false;

        public Boolean Feature17 = false;

        public Boolean Feature18 = false;

        public Boolean Feature19 = false;

        public Boolean Feature20 = false;

        public Boolean Feature21 = false;

        public Boolean Feature22 = false;

        public Boolean Feature23 = false;

        public Boolean Feature24 = false;

        public Boolean Feature25 = false;

        public Boolean Feature26 = false;

        public Boolean Feature27 = false;

        public Boolean Feature28 = false;

        public Boolean Feature29 = false;

        public Boolean Feature30 = false;

        #endregion

        #region Misc

        public Byte FilterCount = 0;

        public Int32 UiLcid
        {
            get
            {
                return (UiLanguage.LCID);
            }
            set
            {
                UiLanguage = CultureInfo.GetCultureInfo(value);
            }
        }

        [XmlIgnore]
        internal CultureInfo UiLanguage;

        public Boolean ExportToCollectionXml = false;

        #endregion

        public DefaultValues()
        {
            UiLanguage = GetUILanguage();
        }

        internal static CultureInfo GetUILanguage()
        {
            String uiCulture = Thread.CurrentThread.CurrentUICulture.Name;

            CultureInfo uiLanguage;
            if (uiCulture.StartsWith("de"))
            {
                uiLanguage = CultureInfo.GetCultureInfo("de");
            }
            else if (uiCulture.StartsWith("fr"))
            {
                uiLanguage = CultureInfo.GetCultureInfo("fr");
            }
            else
            {
                uiLanguage = CultureInfo.GetCultureInfo("en");
            }

            return (uiLanguage);
        }
    }
}