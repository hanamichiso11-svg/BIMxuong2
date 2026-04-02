using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace BIMxuong2
{
    internal static class ModelData
    {
        public const double MmToFt = 0.00328083989501312;

        public const string ProjectName = "Xuong che bien chan ga cay so 2";

        public static readonly double[] GridXmm = { 0, 8000, 16000, 24000, 32000, 40000 };
        public static readonly string[] GridXNames = { "A", "B", "C", "D", "E", "F" };

        public static readonly double[] GridYmm = { 0, 8500, 17000, 25500, 34000, 42500, 51000, 59500, 68000, 76500, 85000 };
        public static readonly string[] GridYNames = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11" };

        public const double BaseLevelMm = 0;
        public const double EaveLevelMm = 7000;
        public const double RidgeLevelMm = 11138;
        public const double BuildingWidthMm = 40000;
        public const double BuildingLengthMm = 85000;
        public const double RidgeXmm = 20000;

        public const double RoofSlope = (RidgeLevelMm - EaveLevelMm) / RidgeXmm;

        public const double RoofPurlinSpacingMm = 1200;
        public const double RoofPurlinEdgeOffsetMm = 300;

        public const double WallHeightMm = 7000;

        public static double FtFromMm(double mm)
        {
            return mm * MmToFt;
        }

        public static XYZ PtMm(double xMm, double yMm, double zMm)
        {
            return new XYZ(FtFromMm(xMm), FtFromMm(yMm), FtFromMm(zMm));
        }

        public static double RoofZmm(double xMm)
        {
            double dx = Math.Abs(xMm - RidgeXmm);
            return EaveLevelMm + RoofSlope * (RidgeXmm - dx);
        }

        public static double ColumnTopOffsetMm(double xMm)
        {
            double roofZ = RoofZmm(xMm);
            return Math.Max(0.0, roofZ - EaveLevelMm);
        }

        public static IList<double> BuildPurlinXmm()
        {
            List<double> xs = new List<double>();
            double spacingPlanMm = RoofPurlinSpacingMm / Math.Sqrt(1 + RoofSlope * RoofSlope);

            AddPurlins(xs, RoofPurlinEdgeOffsetMm, RidgeXmm - RoofPurlinEdgeOffsetMm, spacingPlanMm);
            AddPurlins(xs, RidgeXmm + RoofPurlinEdgeOffsetMm, BuildingWidthMm - RoofPurlinEdgeOffsetMm, spacingPlanMm);

            return xs;
        }

        private static void AddPurlins(List<double> xs, double startMm, double endMm, double spacingMm)
        {
            if (endMm <= startMm) return;

            for (double x = startMm; x <= endMm + 0.1; x += spacingMm)
            {
                if (NearPrimaryFrame(x)) continue;
                if (Math.Abs(x - RidgeXmm) < 250) continue;
                xs.Add(Math.Round(x, 0));
            }
        }

        private static bool NearPrimaryFrame(double xMm)
        {
            for (int i = 0; i < GridXmm.Length; i++)
            {
                if (Math.Abs(GridXmm[i] - xMm) < 250) return true;
            }
            return false;
        }
    }
}
