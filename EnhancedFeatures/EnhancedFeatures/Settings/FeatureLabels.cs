using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace DoenaSoft.DVDProfiler.EnhancedFeatures
{
    internal sealed class FeatureLabels
    {
        private readonly DefaultValues DefaultValues;

        private readonly IEnumerable<FieldInfo> FieldInfos;

        internal FeatureLabels(DefaultValues dv)
        {
            DefaultValues = dv;

            FieldInfos = dv.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);
        }

        internal String this[Byte featureIndex]
        {
            get
            {
                var field = this.GetLabelField(featureIndex);

                var value = (String)(field.GetValue(DefaultValues));

                return (value);
            }
            set
            {
                var field = this.GetLabelField(featureIndex);

                field.SetValue(DefaultValues, value);
            }
        }

        private FieldInfo GetLabelField(Byte featureIndex)
            => (FieldInfos.Where(fi => fi.Name == $"{Constants.Feature}{featureIndex}{Constants.LabelSuffix}").First());
    }
}