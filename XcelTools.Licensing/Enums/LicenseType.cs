using ReInvented.Shared.TypeConverters;

using System.ComponentModel;

namespace XcelTools.Licensing.Enums
{
    [TypeConverter(typeof(EnumToDescriptionTypeConverter))]
    public enum LicenseType
    {
        [Description("Short Trial")]
        ShortTrial = 7,

        [Description("Evaluation")]
        Evaluation = 30,

        [Description("Annual")]
        Annual = 365,

        [Description("Perpetual")]
        Perpetual = 0
    }
}
