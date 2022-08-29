using ReInvented.Shared.TypeConverters;

using System.ComponentModel;

namespace XcelTools.Licensing.Enums
{
    [TypeConverter(typeof(EnumToDescriptionTypeConverter))]
    public enum ApplicationModule
    {
        [Description("Excel AddIn Module")]
        Xtractor,

        [Description("Thickener Design Automation Module")]
        StaadAutomation

    }
}
