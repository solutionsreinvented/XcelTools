using ReInvented.Sections.Domain.Models;

using XcelTools.Xtractor.Models.Sections;

namespace XcelTools.Extensions
{
    public static class RolledSectionOShapeExtensions
    {
        public static SectionProperties RetrieveSpecificProperties(this RolledSectionOShape section, SectionProperties sp)
        {
            sp.OD = section.OD;                                 /* mm */
            sp.ID = section.ID;                                 /* mm */
            sp.Tw = section.Tw;                                 /* mm */

            return sp;
        }
        public static SectionProperties RetrieveSpecificProperties(this RolledSectionHShape section, SectionProperties sp)
        {
            sp.H = section.H;                                   /* mm */
            sp.Bf = section.Bf;                                 /* mm */
            sp.Tw = section.Tw;                                 /* mm */
            sp.Tf = section.Tf;                                 /* mm */
            sp.R1 = section.R1;                                 /* mm */
            sp.R2 = section.R2;                                 /* mm */
            sp.Hi = section.Hi;                                 /* mm */
            sp.D = section.D;                                   /* mm */
            sp.Alpha = section.Alpha;                           /* ° */
            sp.K = (section.H - section.D) / 2;                 /* mm */
            sp.K1 = sp.Tw / 2 + sp.R1;
            //sp.H =  CSng(.SelectSingleNode("h").InnerText) ' mm
            //sp.Bf = CSng(.SelectSingleNode("bf").InnerText) ' mm
            //sp.Tw = CSng(.SelectSingleNode("tw").InnerText) ' mm
            //sp.Tf = CSng(.SelectSingleNode("tf").InnerText) ' mm
            //sp.R1 = CSng(.SelectSingleNode("r1").InnerText) ' mm
            //sp.R2 = CSng(.SelectSingleNode("r2").InnerText) ' mm
            //sp.Hi = CSng(.SelectSingleNode("hi").InnerText) ' mm
            //sp.D = CSng(.SelectSingleNode("d").InnerText) ' mm
            //sp.Alpha = CDbl(.SelectSingleNode("alpha").InnerText) ' Degrees
            //sp.K = CSng((CSng(.SelectSingleNode("h").InnerText) - .SelectSingleNode("d").InnerText) / 2)
            //sp.K1 = SP.tw / 2 + SP.r1
            //sp.Ss = CSng(.SelectSingleNode("ss").InnerText) ' mm

            return sp;
        }
    }
}
