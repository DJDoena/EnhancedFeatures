using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace DoenaSoft.DVDProfiler.EnhancedFeatures
{
    internal sealed class ExcelFeatures
    {
        private readonly DefaultValues DefaultValues;

        private readonly IEnumerable<FieldInfo> FieldInfos;

        internal ExcelFeatures(DefaultValues dv)
        {
            DefaultValues = dv;

            FieldInfos = dv.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);
        }

        internal Boolean this[Byte featureIndex]
        {
            get
            {
                var field = this.GetLabelField(featureIndex);

                var value = (Boolean)(field.GetValue(DefaultValues));

                return (value);
            }
            set
            {
                var field = this.GetLabelField(featureIndex);

                field.SetValue(DefaultValues, value);
            }
        }

        private FieldInfo GetLabelField(Byte featureIndex)
            => (FieldInfos.Where(fi => fi.Name == $"{Constants.Feature}{featureIndex}").First());
    }
}