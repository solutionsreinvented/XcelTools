using ReInvented.Sections.Domain.Interfaces;
using ReInvented.Sections.Domain.Models;

using XcelTools.Xtractor.Models.Sections;

namespace XcelTools.Extensions
{
    public static class IRolledSectionExtensions
    {
        public static SectionProperties RetrieveSpecificProperties(this SectionProperties sp, IRolledSection section)
        {
            if (section is RolledSectionOShape oShape)
            {

            }
            return sp;
        }
    }
}
