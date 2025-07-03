using System;
using System.Collections.Generic;
using System.Linq;

/** 
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/

namespace Nastranh5
{
    class Tensors
    {
        //private object intp;

        public Tensors() { }

        public double ElementAverage(double[] Nodalvalues)
        {
            double avgValue;
            int nodeCount;
            nodeCount = Nodalvalues.Length;
            avgValue = Nodalvalues.Sum() / nodeCount;
            return avgValue;
        }


        public object[] vonMises2DCPLX(Dictionary<string, string> strD)
        {
            double[,] A2D = new double[,] { { 1.0, -0.5, 0 } ,
                                            { -0.5, 1.0, 0.0 } ,
                                            { 0.0, 0.0, 3.0 } };

            // Code begin vM2D
            //Real Part

            double[] Result = CPLXValueArgument(Convert.ToDouble(strD["X1R"]), Convert.ToDouble(strD["X1I"]));
            double s1 = Result[0];
            double a1 = Result[1];

            Result = CPLXValueArgument(Convert.ToDouble(strD["Y1R"]), Convert.ToDouble(strD["Y1R"]));
            double s2 = Result[0];
            double a2 = Result[1];

            Result = CPLXValueArgument(Convert.ToDouble(strD["TXY1R"]), Convert.ToDouble(strD["TXY1I"]));
            double s12 = Result[0];
            double a12 = Result[1];

            double rhs = (1.0 * ((s1 * s1 * Math.Sin(2.0 * a1)) + (s2 * s2 * Math.Sin(2.0 * a2)) +
                       3.0 * (s12 * s12 * Math.Sin(2.0 * a12)) - (Math.Abs(s1) * Math.Abs(s2) * Math.Sin(a1 + a2))) /
                     ((s1 * s1 * Math.Cos(2.0 * a1)) + (s2 * s2 * Math.Cos(2.0 * a2)) +
                       3.0 * (s12 * s12 * Math.Cos(2.0 * a12)) - (Math.Abs(s1) * Math.Abs(s2) * Math.Cos(a1 + a2))));
            double o1 = -0.5 * Math.Atan(rhs);
            double svm1 = Math.Sqrt((s1 * s1 * Math.Cos(a1 + o1) * Math.Cos(a1 + o1)) +
                       (s2 * s2 * Math.Cos(a2 + o1) * Math.Cos(a2 + o1)) +
                       3.0 * (s12 * s12 * Math.Cos(a12 + o1) * Math.Cos(a12 + o1)) -
                       (Math.Abs(s1) * Math.Abs(s2) * Math.Cos(a1 + o1) * Math.Cos(a2 + o1)));
            o1 = o1 + Math.PI / 2.0;
            double svm2 = Math.Sqrt((s1 * s1 * Math.Cos(a1 + o1) * Math.Cos(a1 + o1)) +
                       (s2 * s2 * Math.Cos(a2 + o1) * Math.Cos(a2 + o1)) +
                       3.0 * (s12 * s12 * Math.Cos(a12 + o1) * Math.Cos(a12 + o1)) -
                       (Math.Abs(s1) * Math.Abs(s2) * Math.Cos(a1 + o1) * Math.Cos(a2 + o1)));
            double cvm1 = Math.Max(svm1, svm2);

            //# top

            Result = CPLXValueArgument(Convert.ToDouble(strD["X2R"]), Convert.ToDouble(strD["X2I"]));
            s1 = Result[0];
            a1 = Result[1];

            Result = CPLXValueArgument(Convert.ToDouble(strD["Y2R"]), Convert.ToDouble(strD["Y2R"]));
            s2 = Result[0];
            a2 = Result[1];

            Result = CPLXValueArgument(Convert.ToDouble(strD["TXY2R"]), Convert.ToDouble(strD["TXY2I"]));
            s12 = Result[0];
            a12 = Result[1];

            rhs = (1.0 * ((s1 * s1 * Math.Sin(2.0 * a1)) + (s2 * s2 * Math.Sin(2.0 * a2)) +
                         3.0 * (s12 * s12 * Math.Sin(2.0 * a12)) - (Math.Abs(s1) * Math.Abs(s2) * Math.Sin(a1 + a2))) /
                       ((s1 * s1 * Math.Cos(2.0 * a1)) + (s2 * s2 * Math.Cos(2.0 * a2)) +
                         3.0 * (s12 * s12 * Math.Cos(2.0 * a12)) - (Math.Abs(s1) * Math.Abs(s2) * Math.Cos(a1 + a2))));
            o1 = -0.5 * Math.Atan(rhs);
            svm1 = Math.Sqrt((s1 * s1 * Math.Cos(a1 + o1) * Math.Cos(a1 + o1)) +
                  (s2 * s2 * Math.Cos(a2 + o1) * Math.Cos(a2 + o1)) +
                  3.0 * (s12 * s12 * Math.Cos(a12 + o1) * Math.Cos(a12 + o1)) -
                  (Math.Abs(s1) * Math.Abs(s2) * Math.Cos(a1 + o1) * Math.Cos(a2 + o1)));
            o1 = o1 + Math.PI / 2.0;
            svm2 = Math.Sqrt((s1 * s1 * Math.Cos(a1 + o1) * Math.Cos(a1 + o1)) +
                  (s2 * s2 * Math.Cos(a2 + o1) * Math.Cos(a2 + o1)) +
                  3.0 * (s12 * s12 * Math.Cos(a12 + o1) * Math.Cos(a12 + o1)) -
                  (Math.Abs(s1) * Math.Abs(s2) * Math.Cos(a1 + o1) * Math.Cos(a2 + o1)));

            double cvm2 = Math.Max(svm1, svm2);

            object[] vM = { cvm1, cvm2 };
            return vM;


            //Code end

        }

        public object[] vonMises2D(Dictionary<string, string> strVect)
        {
            double s1 = Convert.ToDouble(strVect["X1"]);
            double s2 = Convert.ToDouble(strVect["Y1"]);
            double s12 = Convert.ToDouble(strVect["XY1"]);


            double vM1 = Math.Sqrt(Math.Pow(s1, 2) + Math.Pow(s2, 2) - s1 * s2 + 3 * Math.Pow(s12, 2));

            s1 = Convert.ToDouble(strVect["X2"]);
            s2 = Convert.ToDouble(strVect["Y2"]);
            s12 = Convert.ToDouble(strVect["XY2"]);

            double vM2 = Math.Sqrt(Math.Pow(s1, 2) + Math.Pow(s2, 2) - s1 * s2 + 3 * Math.Pow(s12, 2));

            object[] vM = { vM1, vM2 };
            return vM;

        }

        private double[] CPLXValueArgument(double RePart, double ImPart)
        {
            double s1 = Math.Sqrt(Math.Pow(RePart, 2) + Math.Pow(ImPart, 2));
            double a1 = Math.Atan(ImPart / RePart);
            if (RePart < 0.0)
            {
                a1 = a1 + Math.PI;
            }

            return new double[] { s1, a1 };
        }


        public object[] vonMises3DCPLX(Dictionary<string, string> strVect)
        {

            double[,] A3D = new double[,] { { 1.0,  -0.5, -0.5,  0, 0, 0},
                                           { -0.5,    1, -0.5,  0, 0, 0},
                                           { -0.5, -0.5,    1,  0, 0, 0},
                                           {  0,      0,    0,  3, 0, 0},
                                           {  0,      0,    0,  0, 3, 0},
                                           {  0,      0,    0,  0, 0, 3}};


            double[] Result = CPLXValueArgument(Convert.ToDouble(strVect["XR"]), Convert.ToDouble(strVect["XI"]));
            double s1 = Result[0];
            double a1 = Result[1];

            Result = CPLXValueArgument(Convert.ToDouble(strVect["YR"]), Convert.ToDouble(strVect["YI"]));
            double s2 = Result[0];
            double a2 = Result[1];

            Result = CPLXValueArgument(Convert.ToDouble(strVect["ZR"]), Convert.ToDouble(strVect["ZI"]));
            double s3 = Result[0];
            double a3 = Result[1];

            Result = CPLXValueArgument(Convert.ToDouble(strVect["TXYR"]), Convert.ToDouble(strVect["TXYI"]));
            double s12 = Result[0];
            double a12 = Result[1];

            Result = CPLXValueArgument(Convert.ToDouble(strVect["TYZR"]), Convert.ToDouble(strVect["TYZI"]));
            double s23 = Result[0];
            double a23 = Result[1];

            Result = CPLXValueArgument(Convert.ToDouble(strVect["TZXR"]), Convert.ToDouble(strVect["TZXI"]));
            double s31 = Result[0];
            double a31 = Result[1];

            //#    
            double rhs = (1.0 * ((s1 * s1 * Math.Sin(2.0 * a1)) + (s2 * s2 * Math.Sin(2.0 * a2)) + (s3 * s3 * Math.Sin(2.0 * a3)) +
                       3.0 * ((s12 * s12 * Math.Sin(2.0 * a12)) + (s23 * s23 * Math.Sin(2.0 * a23)) +
                       (s31 * s31 * Math.Sin(2.0 * a31))) - (Math.Abs(s1) * Math.Abs(s2) * Math.Sin(a1 + a2)) -
                       (Math.Abs(s1) * Math.Abs(s3) * Math.Sin(a1 + a3)) - (Math.Abs(s2) * Math.Abs(s3) * Math.Sin(a2 + a3))) /
                     ((s1 * s1 * Math.Cos(2.0 * a1)) + (s2 * s2 * Math.Cos(2.0 * a2)) + (s3 * s3 * Math.Cos(2.0 * a3)) +
                       3.0 * ((s12 * s12 * Math.Cos(2.0 * a12)) + (s23 * s23 * Math.Cos(2.0 * a23)) +
                       (s31 * s31 * Math.Cos(2.0 * a31))) - (Math.Abs(s1) * Math.Abs(s2) * Math.Cos(a1 + a2)) -
                       (Math.Abs(s1) * Math.Abs(s3) * Math.Cos(a1 + a3)) - (Math.Abs(s2) * Math.Abs(s3) * Math.Cos(a2 + a3))));

            double o1 = -0.5 * Math.Atan(rhs);

            double svm1 = Math.Sqrt((s1 * s1 * Math.Cos(a1 + o1) * Math.Cos(a1 + o1)) +
                       (s2 * s2 * Math.Cos(a2 + o1) * Math.Cos(a2 + o1)) +
                       (s3 * s3 * Math.Cos(a3 + o1) * Math.Cos(a3 + o1)) +
                       3.0 * ((s12 * s12 * Math.Cos(a12 + o1) * Math.Cos(a12 + o1)) +
                            (s23 * s23 * Math.Cos(a23 + o1) * Math.Cos(a23 + o1)) +
                            (s31 * s31 * Math.Cos(a31 + o1) * Math.Cos(a31 + o1))) -
                            (Math.Abs(s1) * Math.Abs(s2) * Math.Cos(a1 + o1) * Math.Cos(a2 + o1)) -
                            (Math.Abs(s1) * Math.Abs(s3) * Math.Cos(a1 + o1) * Math.Cos(a3 + o1)) -
                            (Math.Abs(s2) * Math.Abs(s3) * Math.Cos(a2 + o1) * Math.Cos(a3 + o1)));

            o1 = o1 + Math.PI / 2.0;
            double svm2 = Math.Sqrt((s1 * s1 * Math.Cos(a1 + o1) * Math.Cos(a1 + o1)) +
                       (s2 * s2 * Math.Cos(a2 + o1) * Math.Cos(a2 + o1)) +
                       (s3 * s3 * Math.Cos(a3 + o1) * Math.Cos(a3 + o1)) +
                       3.0 * ((s12 * s12 * Math.Cos(a12 + o1) * Math.Cos(a12 + o1)) +
                            (s23 * s23 * Math.Cos(a23 + o1) * Math.Cos(a23 + o1)) +
                            (s31 * s31 * Math.Cos(a31 + o1) * Math.Cos(a31 + o1))) -
                            (Math.Abs(s1) * Math.Abs(s2) * Math.Cos(a1 + o1) * Math.Cos(a2 + o1)) -
                            (Math.Abs(s1) * Math.Abs(s3) * Math.Cos(a1 + o1) * Math.Cos(a3 + o1)) -
                            (Math.Abs(s2) * Math.Abs(s3) * Math.Cos(a2 + o1) * Math.Cos(a3 + o1)));


            object[] cvm1 = { Math.Max(svm1, svm2) };
            return cvm1;
        }


        public object[] Principal2D(Dictionary<string, string> strVect)
        {
            double s1 = Convert.ToDouble(strVect["X1"]);
            double s2 = Convert.ToDouble(strVect["Y1"]);
            double s12 = Convert.ToDouble(strVect["XY1"]);

            double MaxPrinc1 = (s1 + s2) / 2 + 0.5 * Math.Sqrt(Math.Pow(s1 - s2, 2) + 4 * Math.Pow(s12, 2));
            double MinPrinc1 = (s1 + s2) / 2 - 0.5 * Math.Sqrt(Math.Pow(s1 - s2, 2) + 4 * Math.Pow(s12, 2));
            double angPrinc1 = Math.Atan(2 * s12 / (s1 - s2)) / 2;

            s1 = Convert.ToDouble(strVect["X2"]);
            s2 = Convert.ToDouble(strVect["Y2"]);
            s12 = Convert.ToDouble(strVect["XY2"]);

            double MaxPrinc2 = (s1 + s2) / 2 + 0.5 * Math.Sqrt(Math.Pow(s1 - s2, 2) + 4 * Math.Pow(s12, 2));
            double MinPrinc2 = (s1 + s2) / 2 - 0.5 * Math.Sqrt(Math.Pow(s1 - s2, 2) + 4 * Math.Pow(s12, 2));
            double angPrinc2 = Math.Atan(2 * s12 / (s1 - s2)) / 2;

            //object[] PrinComps = { MaxPrinc1, MinPrinc1, angPrinc1, MaxPrinc2, MinPrinc2, angPrinc2 };
            //return PrinComps;
            return new object[] { MaxPrinc1, MinPrinc1, angPrinc1, MaxPrinc2, MinPrinc2, angPrinc2 };
        }

        public object[] Principal3D(Dictionary<string, string> strVect)
        {
            double s1 = Convert.ToDouble(strVect["X"]);
            double s2 = Convert.ToDouble(strVect["Y"]);
            double s3 = Convert.ToDouble(strVect["Z"]);
            double s12 = Convert.ToDouble(strVect["TXY"]);
            double s23 = Convert.ToDouble(strVect["TYZ"]);
            double s31 = Convert.ToDouble(strVect["TZX"]);
            double phi;
            double I1, I2, I3;

            I1 = s1 + s2 + s3;
            I2 = s1 * s2 + s2 * s3 + s3 * s1 - Math.Pow(s12, 2) - Math.Pow(s23, 2) - Math.Pow(s31, 2);
            I3 = s1 * s2 * s3 - s1 * Math.Pow(s23, 2) - s2 * Math.Pow(s31, 2) - s3 * Math.Pow(s12, 2) + 2 * s12 * s23 * s31;
            phi = 1.0 / 3.0 * Math.Acos((2 * Math.Pow(I1, 3) - 9.0 * I1 * I2 + 27.0 * I3) / (2 * Math.Pow(I1 * I1 - 3 * I2, 3.0 / 2.0)));

            double MaxPrinc = I1 / 3.0 + 2.0 / 3.0 * Math.Sqrt(I1 * I1 - 3 * I2) * Math.Cos(phi);
            double MidPrinc = I1 / 3.0 + 2.0 / 3.0 * Math.Sqrt(I1 * I1 - 3 * I2) * Math.Cos(phi - 2.0 * Math.PI / 3.0);
            double MinPrinc = I1 / 3.0 + 2.0 / 3.0 * Math.Sqrt(I1 * I1 - 3 * I2) * Math.Cos(phi - 4.0 * Math.PI / 3.0);

            object[] PrinComps = { MaxPrinc, MidPrinc, MinPrinc };
            PrinComps = PrinComps.OrderByDescending(c => c).ToArray();
            return PrinComps;
        }

        public object[] PrincipalStrain2D(Dictionary<string, string> strVect)
        {
            double s1 = Convert.ToDouble(strVect["X1"]);
            double s2 = Convert.ToDouble(strVect["Y1"]);
            double s12 = Convert.ToDouble(strVect["XY1"]);

            double MaxPrinc1 = (s1 + s2) / 2 + 0.5 * Math.Sqrt(Math.Pow(s1 - s2, 2) + Math.Pow(s12, 2));
            double MinPrinc1 = (s1 + s2) / 2 - 0.5 * Math.Sqrt(Math.Pow(s1 - s2, 2) + Math.Pow(s12, 2));
            double angPrinc1 = Math.Atan(2 * s12 / (s1 - s2)) / 2;

            s1 = Convert.ToDouble(strVect["X2"]);
            s2 = Convert.ToDouble(strVect["Y2"]);
            s12 = Convert.ToDouble(strVect["XY2"]);

            double MaxPrinc2 = (s1 + s2) / 2 + 0.5 * Math.Sqrt(Math.Pow(s1 - s2, 2) + Math.Pow(s12, 2));
            double MinPrinc2 = (s1 + s2) / 2 - 0.5 * Math.Sqrt(Math.Pow(s1 - s2, 2) + Math.Pow(s12, 2));
            double angPrinc2 = Math.Atan(2 * s12 / (s1 - s2)) / 2;

            //object[] PrinComps = { MaxPrinc1, MinPrinc1, angPrinc1, MaxPrinc2, MinPrinc2, angPrinc2 };
            //return PrinComps;
            return new object[] { MaxPrinc1, MinPrinc1, angPrinc1, MaxPrinc2, MinPrinc2, angPrinc2 };
        }

        public object[] PrincipalStrain3D(Dictionary<string, string> strVect)
        {
            double s1 = Convert.ToDouble(strVect["X"]);
            double s2 = Convert.ToDouble(strVect["Y"]);
            double s3 = Convert.ToDouble(strVect["Z"]);
            double s12 = Convert.ToDouble(strVect["TXY"]);
            double s23 = Convert.ToDouble(strVect["TYZ"]);
            double s31 = Convert.ToDouble(strVect["TZX"]);
            double phi;
            double I1, I2, I3;

            I1 = s1 + s2 + s3;
            I2 = s1 * s2 + s2 * s3 + s3 * s1 - Math.Pow(s12, 2) - Math.Pow(s23, 2) - Math.Pow(s31, 2);
            I3 = s1 * s2 * s3 - s1 * Math.Pow(s23, 2) - s2 * Math.Pow(s31, 2) - s3 * Math.Pow(s12, 2) + 2 * s12 * s23 * s31;
            //phi = 1.0 / 3.0 * Math.Acos((2 * Math.Pow(I1, 3) - 9.0 * I1 * I2 + 27.0 * I3) / (2 * Math.Pow(I1 * I1 - 3 * I2, 3.0 / 2.0)));
            double Q = (3 * I2 - Math.Pow(I1, 2)) / 9;
            double R = (2 * Math.Pow(I1, 3) - 9.0 * I1 * I2 + 27.0 * I3) / 54.0;

            phi = Math.Acos(R / Math.Sqrt(-1 * Math.Pow(Q, 3)));

            double MaxPrinc = I1 / 3.0 + 2.0 * Math.Sqrt(-Q) * Math.Cos(phi / 3);
            double MidPrinc = I1 / 3.0 + 2.0 * Math.Sqrt(-Q) * Math.Cos((phi + 2.0 * Math.PI) / 3);
            double MinPrinc = I1 / 3.0 + 2.0 * Math.Sqrt(-Q) * Math.Cos((phi + 4.0 * Math.PI) / 3);

            object[] PrinComps = { MaxPrinc, MidPrinc, MinPrinc };
            PrinComps = PrinComps.OrderByDescending(c => c).ToArray();
            return PrinComps;
        }

        public object[] vonMises3D(Dictionary<string, string> strVect)
        {
            double s1 = Convert.ToDouble(strVect["X"]);
            double s2 = Convert.ToDouble(strVect["Y"]);
            double s3 = Convert.ToDouble(strVect["Z"]);
            double s12 = Convert.ToDouble(strVect["TXY"]);
            double s23 = Convert.ToDouble(strVect["TYZ"]);
            double s31 = Convert.ToDouble(strVect["TZX"]);


            double vM1 = Math.Sqrt(Math.Pow(s1, 2) + Math.Pow(s2, 2) + Math.Pow(s3, 2)
                - s1 * s2 - s2 * s3 - s3 * s1 + 3 * (Math.Pow(s12, 2) + Math.Pow(s23, 2) + Math.Pow(s31, 2)));


            object[] vM = { vM1 };
            return vM;
        }

        public object[] vonMisesStrain3D(Dictionary<string, string> strVect)
        {
            double s1 = Convert.ToDouble(strVect["X"]);
            double s2 = Convert.ToDouble(strVect["Y"]);
            double s3 = Convert.ToDouble(strVect["Z"]);
            double s12 = Convert.ToDouble(strVect["TXY"]);
            double s23 = Convert.ToDouble(strVect["TYZ"]);
            double s31 = Convert.ToDouble(strVect["TZX"]);


            double vM1 = Math.Sqrt(2.0 / 9.0 * (Math.Pow((s1 - s2), 2) + Math.Pow((s2 - s3), 2) + Math.Pow((s3 - s1), 2))
                 + 1.0 / 3.0 * (Math.Pow(s12 / 2, 2) + Math.Pow(s23 / 2, 2) + Math.Pow(s31 / 2, 2)));


            object[] vM = { vM1 };
            return vM;
        }



    }
}
