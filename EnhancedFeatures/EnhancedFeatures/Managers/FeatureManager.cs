using System;
using Invelos.DVDProfilerPlugin;

namespace DoenaSoft.DVDProfiler.EnhancedFeatures
{
    internal sealed class FeatureManager
    {
        private readonly IDVDInfo Profile;

        internal FeatureManager(IDVDInfo profile)
        {
            Profile = profile;
        }

        internal Boolean GetFeature(Byte index)
            => (Profile.GetCustomBool(Constants.FieldDomain, $"Feature{index}", Constants.ReadKey, false));

        internal void SetFeature(Byte index
            , Boolean value)
        {
            Profile.SetCustomBool(Constants.FieldDomain, $"Feature{index}", InternalConstants.WriteKey, value);
        }
    }
}