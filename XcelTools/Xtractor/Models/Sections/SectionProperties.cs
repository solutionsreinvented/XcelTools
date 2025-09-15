using ReInvented.Sections.Domain.Interfaces;
using ReInvented.Sections.Domain.Models;
using ReInvented.Shared;

namespace XcelTools.Xtractor.Models.Sections
{
    public struct SectionProperties
    {
        #region Public Properties

        public string Designation;
        public double Mass;
        public double MassFps;

        #region Pipe Specific Properties

        public double OD;
        public double ID;

        #endregion

        public double H;
        public double Bf;
        public double Tw;
        public double Tf;
        public double R1;
        public double R2;
        public double A;
        public double Hi;
        public double D;
        public double Phi;
        public double Pmin;
        public double Pmax;
        public double ALo;
        public double ALi;
        public double AGo;
        public double AGi;
        public double Iy;
        public double Wely;
        public double Wply;
        public double Ry;
        public double Avz;
        public double Iz;
        public double Welz;
        public double Wplz;
        public double Rz;
        public double Avy;
        public double Ss;
        public double It;
        public double Iw;
        public double K;
        public double K1;
        public double Alpha;
        public double Cy;
        public double Ey;
        public double Cz;
        public double Ez;
        public double Iu;
        public double Iv;
        public double Ru;
        public double Rv;

        #endregion

        #region Public Static Methods

        public static SectionProperties GetFrom(IRolledSection section)
        {
            double Ten = 10.0;
            ///Note: Units marked next to the properties are the units in which the data is stored in the database.
            SectionProperties sp = new SectionProperties()
            {
                Designation = section.Designation,
                Mass = section.Mass,                                    /* kg/m */
                MassFps = section.MassFps,                              /* pounds/foot */
                Iy = section.Iy * Ten.RestTo(4),                        /* cm⁴ */
                Wely = section.Wely * Ten.Cube(),                       /* cm³ */
                Wply = section.Wply * Ten.Cube(),                       /* cm³ */
                Ry = section.Ry * Ten,                                  /* cm³ */
                Avy = section.Avy * Ten.Square(),                       /* cm² */
                Iz = section.Iz * Ten.RestTo(4),                        /* cm⁴ */
                Welz = section.Welz * Ten.Cube(),                       /* cm³ */
                Wplz = section.Wplz * Ten.Cube(),                       /* cm³ */
                Rz = section.Rz * Ten,                                  /* cm */
                Avz = section.Avz * Ten.Square(),                       /* cm² */
                It = section.It * Ten.RestTo(4),                        /* cm⁴ */
                Iw = section.Iw * Ten.Cube() * Ten.RestTo(6),           /* x10³ cm⁶ */
                Cy = section.Cy,                                        /* mm */
                Cz = section.Cz,                                        /* mm */
                Ey = section.Ey,                                        /* mm */
                Ez = section.Ez,                                        /* mm */
                A = section.A * Ten.Square(),                           /* cm² */

                ALo = section.ALO,                                      /* m²/m */
                AGo = section.AGO,                                      /* m²/t */
                Iu = section.Iu * Ten.RestTo(4),                        /* cm⁴ */
                Iv = section.Iu * Ten.RestTo(4),                        /* cm⁴ */
                Ru = section.Ru * Ten,                                  /* cm */
                Rv = section.Rv * Ten,                                  /* cm */
            };

            if (section is RolledSectionOShape pipeSection)
            {
                sp.OD = pipeSection.OD;                                 /* mm */
                sp.ID = pipeSection.ID;                                 /* mm */
                sp.Tw = pipeSection.Tw;                                 /* mm */
            }
            else
            {
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
            }

            return sp;
        }

        #endregion
    }
}
