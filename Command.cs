// ==============================================================================
// FILE    : Command.cs  (Version 5.1 – BIM UPGRADE (TCVN))
// PROJECT : BIMxuong2
// TARGET  : Revit 2024 | .NET Framework 4.8 | C# 7.3
// TIEU CHUAN: TCVN 5575:2024 | TCVN 2737:2023
// DON VI  : m, mm, kN (Viet Nam) – Revit internal = feet (tu chuyen doi)
//
// SU DUNG: Tao project tu template HUTECH (.rte), sau do chay plugin nay.
//
// NANG CAP V5.0 SO VOI V4.1:
//   [BIM-01] Phan nhom BIM: KC, KCP, LK, MG, SN, TG, CUA, MEP, MAI
//   [BIM-02] Lien ket dung GenericModel + naming LK_ (DirectShape ko ho tro StructConnections)

//   [BIM-03] Tao 3D view + plan view tu dong
//   [BIM-04] Tao schedule cot, dam tu dong
//   [BIM-05] Toi uu sag rod (2 duong/mai/bay thay vi hang tram doan)
//   [BIM-06] Chuan hoa naming toan bo theo nhom BIM
//   [BIM-07] Enhanced validate: dem theo nhom, kiem tra naming
//   [BIM-08] Report chi tiet theo nhom BIM
//
// FIX V4.1 (GIU NGUYEN):
//   [FIX-01] Bien 'g' trung scope trong FindOrCreateGrid (lambda vs local)
//   [FIX-02] Doi ten bien tranh shadow (cc -> ptC trong CreateRoof)
//   [FIX-03] Null-check cho tat ca FamilySymbol/Type truoc khi dung
//   [FIX-04] Them try-catch cho moi buoc trong Build (khong rollback toan bo)
//   [FIX-05] Fix SetCurveInView -> dung Curve property thay vi View-based
//   [FIX-06] Them Project Information (ten cong trinh, dia chi, ma so)
//   [FIX-07] Them Phase mapping cho template
//   [FIX-08] Cai thien report voi chi tiet loi
//   [FIX-09] Them kiem tra trung lap Level/Grid truoc khi tao
//   [FIX-10] Rename tat ca bien de tranh conflict C# 7.3
//
// 23 BUOC BUILD:
//   0.ProjectInfo 1.CleanTemplate 2.Levels 3.Grids
//   4.Symbols 5.Types 6.Foundations 7.Columns 8.Beams
//   9.Purlins 10.SagRods 11.Bracing 12.Connections
//   13.Floors 14.Walls 15.Roof 16.Doors 17.MEP
//   18.Gutter 19.BimViews 20.BimSchedules 21.Cleanup 22.Validate
// ==============================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

namespace BIMxuong2
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            try
            {
                XuongBuilder builder = new XuongBuilder(doc);
                builder.Build();
                TaskDialog.Show("Hoan thanh",
                    "BIM Xuong So 2 – V5.1 BIM\n" +
                    "Template: HUTECH Ket Cau\n" +
                    "Tieu chuan: TCVN 5575:2024 | TCVN 2737:2023\n\n" +
                    builder.GetReport());
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Loi", ex.ToString());
                return Result.Failed;
            }
        }
    }

    // ==========================================================================
    // HANG SO – TAT CA THEO MET (VN)
    // ==========================================================================
    internal static class K
    {
        // [FIX-10] Doi ten class tu 'C' thanh 'K' de tranh shadow voi bien 'c' local
        public const double M2F = 3.280839895; // 1m = 3.2808 feet

        // --- THONG TIN DU AN ---
        public const string PROJ_NAME = "Xuong che bien chan ga cay so 2";
        public const string PROJ_NUM = "XCG-02-2025";
        public const string PROJ_ADDR = "KCN HUTECH, TP.HCM";
        public const string PROJ_ORG = "HUTECH University";

        // --- LUOI TRUC ---
        public static readonly double[] GX = { 0, 8, 16, 24, 32, 40 };        // A-F, buoc 8m
        public static readonly string[] GXN = { "A", "B", "C", "D", "E", "F" };
        public static readonly double[] GY = { 0, 8.5, 17, 25.5, 34, 42.5, 51, 59.5, 68, 76.5, 85 }; // 1-11
        public static readonly string[] GYN = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11" };

        // --- CAO DO (m) ---
        public const double Z_MONG = -0.900;  // day mong BTCT
        public const double Z_DINH_CO = -0.050;  // dinh co mong (sat nen)
        public const double Z_DINH = -0.200;  // mat dat tu nhien / day ranh
        public const double Z_NEN = 0.000;  // nen hoan thien
        public const double Z_LUNG = 3.500;
        public const double Z_EAVE = 7.000;  // dinh cot = mep mai
        public const double Z_RIDGE = 11.138;  // dinh noc

        // --- MAI ---
        public const double SLOPE = 0.2069;  // (11.138-7.000)/20
        public const double W_TOTAL = 40.0;    // rong nha (m)
        public const double L_TOTAL = 85.0;    // dai nha (m)
        public const double RIDGE_X = 20.0;    // noc tai X=20m
        public const double PURLIN_SP = 1.4;     // buoc xa go mai 1400mm

        // --- XA GO TUONG ---
        public static readonly double[] GIRT_Z = { 1.4, 2.8, 4.2, 5.6 };

        // --- NHIP GIANG ---
        public static readonly int[] BRACE_BAY = { 0, 4, 9 };

        // --- COT C1 bien (A,F): H(350-1020)x214x10x10 ---
        public const double COL_H_TOP = 0.350;
        public const double COL_H_BOT = 1.020;
        public const double COL_B = 0.214;
        public const double COL_TW = 0.010;

        // --- COT C2 giua (B-E): H(450-1080)x214x10x12 ---
        public const double COL2_H_TOP = 0.450;
        public const double COL2_H_BOT = 1.080;

        // --- COT HOI CH: H320x186x10x12 ---
        public const double COLH_H = 0.320;
        public const double COLH_B = 0.186;

        // --- KEO K1.1: H(1020-570)x214x10x10 ---
        public const double RAF_H_ROOT = 1.020;
        public const double RAF_H_RIDGE = 0.570;
        public const double RAF_B = 0.214;

        // --- BAN DE COT (TCVN 5575:2024 §14) ---
        public const double BP_L = 0.650;   // 650mm
        public const double BP_W = 0.550;   // 550mm
        public const double BP_T = 0.030;   // 30mm
        public const double BP_STF_H = 0.200;
        public const double BP_STF_T = 0.016;
        public const double GROUT_T = 0.050;   // vua grout 50mm

        // --- BU LONG NEO M30 (TCVN 5575:2024 §14.2) ---
        public const double AB_D = 0.030;
        public const double AB_L = 0.700;  // L=700mm
        public const double AB_PL = 0.500;  // khoang cach doc
        public const double AB_PW = 0.420;  // khoang cach ngang

        // --- END PLATE M20 ---
        public const double EP_H = 0.400;
        public const double EP_W = 0.200;
        public const double EP_T = 0.020;
        public const double BOLT_M20 = 0.020;
        public const double BOLT_L = 0.080;

        // --- FIN PLATE M16 ---
        public const double FP_H = 0.250;
        public const double FP_W = 0.120;
        public const double FP_T = 0.012;
        public const double BOLT_M16 = 0.016;

        // --- GUSSET (TCVN 5575:2024 §14.4) ---
        public const double GS_H = 0.600;
        public const double GS_W = 0.450;
        public const double GS_T = 0.020;

        // --- CLIP ANGLE M12 ---
        public const double CA_L = 0.150;
        public const double CA_W = 0.100;
        public const double CA_T = 0.010;
        public const double BOLT_M12 = 0.012;

        // --- HAUNCH ---
        public const double HAUNCH_L = 1.500;
        public const double HAUNCH_D = 0.500;

        // --- RIDGE SPLICE ---
        public const double RS_W = 0.350;
        public const double RS_H = 0.300;
        public const double RS_T = 0.020;

        // --- MONG BTCT ---
        public const double FD_L = 2.000;
        public const double FD_W = 2.000;
        public const double FD_H = 0.500;
        public const double FD_COL_W = 0.400;
        public const double FD_COL_H = 0.400;
        public const double FD_LOT_T = 0.100;

        // --- MANG NUOC MAI ---
        public const double GUTTER_W = 0.300;
        public const double GUTTER_H = 0.150;
        public const double DOWNPIPE_D = 0.114;

        // --- MEP ---
        public const double LIGHT_SZ = 0.600;
        public const double PIPE_D = 0.110;

        // --- NHOM BIM (V5.0) ---
        public const string G_KC = "KC";   // Ket cau chinh (cot, keo, dam doc)
        public const string G_KCP = "KCP";  // Ket cau phu (xa go, giang, sag rod)
        public const string G_LK = "LK";   // Lien ket (ban de, bu long, gusset...)
        public const string G_MG = "MG";   // Mong (dai, lot, co)
        public const string G_SN = "SN";   // San nen & ranh
        public const string G_TG = "TG";   // Tuong
        public const string G_CUA = "CUA";  // Cua
        public const string G_MEP = "MEP";  // Co dien
        public const string G_MAI = "MAI";  // Mai, mang, ong xoi

        // --- CAO DO DU AN – map level ---
        public static readonly double[] PROJ_ELEVS =
            { Z_MONG, Z_DINH_CO, Z_DINH, Z_NEN, Z_LUNG, Z_EAVE, Z_RIDGE };
        public static readonly string[] PROJ_LV_NAMES =
            { "L00_DayMong_-0.900", "L01_DinhCoMong_-0.050",
              "L02_MatDat_-0.200", "L03_Nen_+0.000",
              "L04_SanLung_+3.500", "L05_Eave_+7.000", "L06_Ridge_+11.138" };

        // === TIEN ICH ===
        public static double ft(double m) { return m * M2F; }
        public static XYZ pt(double xM, double yM, double zM)
        { return new XYZ(xM * M2F, yM * M2F, zM * M2F); }

        public static double RoofZ(double xM)
        {
            double dist = Math.Abs(xM - RIDGE_X);
            return Z_EAVE + SLOPE * (RIDGE_X - dist);
        }

        public static string LX(double x)
        {
            for (int i = 0; i < GX.Length; i++)
                if (Math.Abs(GX[i] - x) < 0.01) return GXN[i];
            return "X" + x.ToString("F1");
        }

        public static string LY(double y)
        {
            for (int i = 0; i < GY.Length; i++)
                if (Math.Abs(GY[i] - y) < 0.01) return GYN[i];
            return "Y" + y.ToString("F1");
        }
    }

    // ==========================================================================
    // BUILDER CHINH
    // ==========================================================================
    internal class XuongBuilder
    {
        private readonly Document _doc;
        private Level _lvMong, _lvDinh, _lvNen, _lvLung, _lvEave, _lvRidge;
        private FamilySymbol _symColBien, _symColGiua, _symColHoi;
        private FamilySymbol _symRafter, _symLong, _symPurlin, _symGirt, _symBrace;
        private FamilySymbol _symDoor, _symDoorDbl, _symDoorRoll, _symWindow;
        private WallType _wtExt, _wtInt, _wtBrick;
        private FloorType _ftNen, _ftRanh;
        private RoofType _roofT;

        private readonly List<FamilyInstance> _cols = new List<FamilyInstance>();
        private readonly List<FamilyInstance> _rafters = new List<FamilyInstance>();
        private readonly List<FamilyInstance> _purlins = new List<FamilyInstance>();
        private int _nCol, _nRaf, _nPur, _nGirt, _nBrace, _nConn;
        private int _nWall, _nDoor, _nFnd, _nMep, _nClean;
        private int _nTmplLvReuse, _nTmplGrReuse, _nTmplDel;
        private int _nView, _nSched, _nSagRod;
        // DirectShape KHONG ho tro OST_StructConnections -> dung GenericModel + naming LK_
        private BuiltInCategory _catConn = BuiltInCategory.OST_GenericModel;
        private readonly List<string> _warnings = new List<string>();
        private readonly List<string> _errors = new List<string>();

        public XuongBuilder(Document doc) { _doc = doc; }

        public string GetReport()
        {
            string warn = _warnings.Count > 0
                ? "\n\nCHI TIET CANH BAO:\n- " + string.Join("\n- ", _warnings.Take(20))
                : "";
            string err = _errors.Count > 0
                ? "\n\nLOI:\n- " + string.Join("\n- ", _errors.Take(10))
                : "";
            return string.Format(
                "=== BIM V5.1 – XUONG SO 2 ===\n" +
                "TIEU CHUAN: TCVN 5575:2024 | TCVN 2737:2023\n\n" +
                "--- TEMPLATE ---\n" +
                "Level reuse: {0} | Grid reuse: {1} | Xoa: {2}\n\n" +
                "--- KET CAU CHINH (KC) ---\n" +
                "Cot:      {3}\n" +
                "Keo/Dam:  {4}\n\n" +
                "--- KET CAU PHU (KCP) ---\n" +
                "Xa go mai:   {5}\n" +
                "Xa go tuong: {6}\n" +
                "Giang:       {7}\n" +
                "Sag rod:     {8}\n\n" +
                "--- LIEN KET (LK) ---\n" +
                "Lien ket: {9}  [GenericModel + LK_]\n\n" +
                "--- MONG (MG) ---\n" +
                "Cum mong: {10}\n\n" +
                "--- TUONG & CUA ---\n" +
                "Tuong: {11} | Cua: {12}\n\n" +
                "--- MEP ---\n" +
                "Thiet bi: {13}\n\n" +
                "--- QUAN LY BIM ---\n" +
                "View: {14} | Schedule: {15}\n" +
                "Cleanup: {16}\n\n" +
                "Canh bao: {17} | Loi: {18}" +
                warn + err,
                _nTmplLvReuse, _nTmplGrReuse, _nTmplDel,
                _nCol, _nRaf,
                _nPur, _nGirt, _nBrace, _nSagRod,
                _nConn,
                _nFnd,
                _nWall, _nDoor,
                _nMep,
                _nView, _nSched,
                _nClean,
                _warnings.Count, _errors.Count);
        }

        // ======================================================================
        // BUILD – 21 buoc
        // [FIX-04] Moi buoc co try-catch rieng, khong rollback toan bo
        // ======================================================================
        public void Build()
        {
            Transaction tx = new Transaction(_doc, "BIM Xuong So 2 V5.1 BIM TCVN");
            tx.Start();
            try
            {
                RunStep(0, "ProjectInfo", SetProjectInfo);
                RunStep(1, "CleanTemplate", CleanTemplateDefaults);
                RunStep(2, "Levels", SetupLevels);
                RunStep(3, "Grids", SetupGrids);
                RunStep(4, "Symbols", ResolveSymbols);
                RunStep(5, "Types", ResolveTypes);
                RunStep(6, "Foundations", CreateFoundations);
                RunStep(7, "Columns", CreateColumns);
                RunStep(8, "Beams", CreateBeams);
                RunStep(9, "Purlins", CreatePurlins);
                RunStep(10, "SagRods", CreateSagRods);
                RunStep(11, "Bracing", CreateBracing);
                RunStep(12, "Connections", CreateConnections);
                RunStep(13, "Floors", CreateFloors);
                RunStep(14, "Walls", CreateWalls);
                RunStep(15, "Roof", CreateRoof);
                RunStep(16, "Doors", PlaceDoors);
                RunStep(17, "MEP", CreateMEP);
                RunStep(18, "Stairs", CreateStairs);
                RunStep(19, "Windows", PlaceWindows);
                RunStep(20, "Gutter", CreateGutter);
                RunStep(21, "BimViews", CreateBimViews);
                RunStep(20, "BimSchedules", CreateBimSchedules);
                RunStep(21, "Cleanup", CleanupModel);
                RunStep(22, "Validate", ValidateModel);
                tx.Commit();
            }
            catch (Exception)
            {
                if (tx.HasStarted()) tx.RollBack();
                throw;
            }
        }

        // [FIX-04] Wrapper de bao ve moi buoc
        private void RunStep(int num, string name, Action action)
        {
            System.Diagnostics.Debug.WriteLine(
                string.Format("[BIM] {0:D2} {1} ...", num, name));
            try
            {
                action();
                System.Diagnostics.Debug.WriteLine(
                    string.Format("[BIM] {0:D2} {1} OK", num, name));
            }
            catch (Exception ex)
            {
                string msg = string.Format("Step {0} ({1}): {2}", num, name, ex.Message);
                _errors.Add(msg);
                System.Diagnostics.Debug.WriteLine("[BIM] ERROR: " + msg);
                // KHONG throw – tiep tuc buoc tiep theo
            }
        }

        private void Warn(string msg) { _warnings.Add(msg); }

        // ======================================================================
        // 0. PROJECT INFORMATION [V4.1-06]
        // ======================================================================
        private void SetProjectInfo()
        {
            try
            {
                ProjectInfo pi = _doc.ProjectInformation;
                if (pi != null)
                {
                    pi.Name = K.PROJ_NAME;
                    pi.Number = K.PROJ_NUM;
                    pi.Address = K.PROJ_ADDR;
                    pi.OrganizationName = K.PROJ_ORG;
                    pi.ClientName = K.PROJ_ORG;
                    pi.Status = "THIET KE";
                    pi.IssueDate = DateTime.Now.ToString("yyyy-MM-dd");

                    // Ghi ghi chu tieu chuan
                    Parameter pComm = pi.get_Parameter(
                        BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                    if (pComm != null && !pComm.IsReadOnly)
                        pComm.Set("TCVN 5575:2024 | TCVN 2737:2023 | BIM V5.1");
                }
            }
            catch (Exception ex)
            {
                Warn("ProjectInfo: " + ex.Message);
            }
        }

        // ======================================================================
        // 1. CLEAN TEMPLATE DEFAULTS
        // ======================================================================
        private void CleanTemplateDefaults()
        {
            List<Level> existLv = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level)).Cast<Level>().ToList();

            HashSet<ElementId> keepLvIds = new HashSet<ElementId>();
            foreach (Level lv in existLv)
            {
                double elev = lv.Elevation / K.M2F;
                foreach (double pe in K.PROJ_ELEVS)
                {
                    if (Math.Abs(elev - pe) < 0.005)
                    {
                        keepLvIds.Add(lv.Id);
                        break;
                    }
                }
            }

            List<Grid> existGr = new FilteredElementCollector(_doc)
                .OfClass(typeof(Grid)).Cast<Grid>().ToList();

            HashSet<ElementId> keepGrIds = new HashSet<ElementId>();
            HashSet<string> projGridNames = new HashSet<string>();
            foreach (string nm in K.GXN) projGridNames.Add(nm);
            foreach (string nm in K.GYN) projGridNames.Add(nm);

            foreach (Grid grd in existGr)
            {
                if (projGridNames.Contains(grd.Name))
                    keepGrIds.Add(grd.Id);
            }

            // Xoa grid thua
            foreach (Grid grd in existGr)
            {
                if (!keepGrIds.Contains(grd.Id))
                {
                    try { _doc.Delete(grd.Id); _nTmplDel++; }
                    catch { Warn("Template: khong the xoa grid " + grd.Name); }
                }
            }

            // Xoa level thua (chi xoa neu khong co element reference)
            foreach (Level lv in existLv)
            {
                if (!keepLvIds.Contains(lv.Id))
                {
                    if (!HasElementsOnLevel(lv.Id))
                    {
                        try { _doc.Delete(lv.Id); _nTmplDel++; }
                        catch { Warn("Template: khong the xoa level " + lv.Name); }
                    }
                    else
                    {
                        Warn("Template: giu level '" + lv.Name + "' (co element reference)");
                    }
                }
            }
        }

        private bool HasElementsOnLevel(ElementId lvId)
        {
            BuiltInCategory[] cats = {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFraming
            };
            foreach (BuiltInCategory cat in cats)
            {
                try
                {
                    FilteredElementCollector col = new FilteredElementCollector(_doc)
                        .OfCategory(cat).WhereElementIsNotElementType();
                    foreach (Element el in col)
                    {
                        Parameter pRef = el.get_Parameter(
                            BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                        if (pRef != null && pRef.AsElementId() == lvId) return true;
                        Parameter pLv = el.get_Parameter(
                            BuiltInParameter.FAMILY_LEVEL_PARAM);
                        if (pLv != null && pLv.AsElementId() == lvId) return true;
                    }
                }
                catch { }
            }
            return false;
        }

        // ======================================================================
        // 2. LEVELS – Tim level co san trong template truoc
        // ======================================================================
        private void SetupLevels()
        {
            _lvMong = FindOrMakeLevel(K.Z_MONG, "L00_DayMong_-0.900");
            FindOrMakeLevel(K.Z_DINH_CO, "L01_DinhCoMong_-0.050");
            _lvDinh = FindOrMakeLevel(K.Z_DINH, "L02_MatDat_-0.200");
            _lvNen = FindOrMakeLevel(K.Z_NEN, "L03_Nen_+0.000");
            _lvLung = FindOrMakeLevel(K.Z_LUNG, "L04_SanLung_+3.500");
            _lvEave = FindOrMakeLevel(K.Z_EAVE, "L05_Eave_+7.000");
            _lvRidge = FindOrMakeLevel(K.Z_RIDGE, "L06_Ridge_+11.138");
        }

        private Level FindOrMakeLevel(double zM, string name)
        {
            double zFt = K.ft(zM);
            List<Level> all = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level)).Cast<Level>().ToList();

            // Tim level co cao do gan nhat (sai so < 5mm)
            Level best = null;
            double bestDist = double.MaxValue;
            foreach (Level lv in all)
            {
                double diff = Math.Abs(lv.Elevation - zFt);
                if (diff < bestDist) { bestDist = diff; best = lv; }
            }

            if (best != null && bestDist < K.ft(0.005))
            {
                // Tai su dung level tu template
                TryRenameLevel(best, name);
                if (Math.Abs(best.Elevation - zFt) > 1e-6)
                    best.Elevation = zFt;
                _nTmplLvReuse++;
                return best;
            }

            // Tim theo ten
            Level byName = all.FirstOrDefault(lv =>
                lv.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (byName != null)
            {
                byName.Elevation = zFt;
                _nTmplLvReuse++;
                return byName;
            }

            // [FIX-09] Kiem tra ten da ton tai chua truoc khi tao
            Level newLv = Level.Create(_doc, zFt);
            TryRenameLevel(newLv, name);
            return newLv;
        }

        private void TryRenameLevel(Level lv, string name)
        {
            if (lv.Name == name) return;
            try { lv.Name = name; }
            catch
            {
                // Ten bi trung – them suffix
                try { lv.Name = name + "_PRJ"; }
                catch { Warn("Level rename failed: " + name); }
            }
        }

        // ======================================================================
        // 3. GRIDS – Tim grid co san trong template truoc
        // [FIX-01] Doi ten bien 'g' thanh 'newGrid' de tranh conflict scope
        // [FIX-05] Bo SetCurveInView, dung Curve property truc tiep
        // ======================================================================
        private void SetupGrids()
        {
            double yMin = K.ft(-3);
            double yMax = K.ft(K.L_TOTAL + 3);
            double xMin = K.ft(-3);
            double xMax = K.ft(K.W_TOTAL + 3);

            // Grid doc (X direction): A-F
            for (int i = 0; i < K.GX.Length; i++)
            {
                double xFt = K.ft(K.GX[i]);
                Line ln = Line.CreateBound(
                    new XYZ(xFt, yMin, 0), new XYZ(xFt, yMax, 0));
                FindOrMakeGrid(K.GXN[i], ln);
            }

            // Grid ngang (Y direction): 1-11
            for (int i = 0; i < K.GY.Length; i++)
            {
                double yFt = K.ft(K.GY[i]);
                Line ln = Line.CreateBound(
                    new XYZ(xMin, yFt, 0), new XYZ(xMax, yFt, 0));
                FindOrMakeGrid(K.GYN[i], ln);
            }
        }

        // [FIX-01] FIXED: bien 'newGrid' thay vi 'g' (tranh conflict voi lambda)
        private Grid FindOrMakeGrid(string name, Line newLine)
        {
            // Tim grid co ten trung trong template
            List<Grid> allGrids = new FilteredElementCollector(_doc)
                .OfClass(typeof(Grid)).Cast<Grid>().ToList();

            Grid exist = allGrids.FirstOrDefault(
                gr => gr.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

            if (exist != null)
            {
                _nTmplGrReuse++;
                return exist;
            }

            // [FIX-09] Kiem tra ten bi trung
            Grid newGrid = Grid.Create(_doc, newLine);
            try { newGrid.Name = name; }
            catch
            {
                try { newGrid.Name = name + "_PRJ"; }
                catch { Warn("Grid rename failed: " + name); }
            }
            return newGrid;
        }

        // ======================================================================
        // 4. SYMBOLS – Uu tien family VN tu template HUTECH
        // ======================================================================
        private void ResolveSymbols()
        {
            // COT BIEN C1
            _symColBien = FindSymbol(BuiltInCategory.OST_StructuralColumns,
                "THEP_H", "H-SECTION", "THEP_HINH_H", "I-SECTION",
                "W21X68", "W21X48", "W21", "Wide Flange", "Column",
                "UC", "HEB", "H");

            // COT GIUA C2
            _symColGiua = FindSymbol(BuiltInCategory.OST_StructuralColumns,
                "THEP_H", "H-SECTION", "W24X76", "W24X68", "W24",
                "UC", "HEB") ?? _symColBien;

            // COT HOI CH
            _symColHoi = FindSymbol(BuiltInCategory.OST_StructuralColumns,
                "THEP_H", "H-SECTION", "W14X38", "W14X30", "W14", "W12",
                "UC", "HEA") ?? _symColBien;

            // KEO
            _symRafter = FindSymbol(BuiltInCategory.OST_StructuralFraming,
                "THEP_H", "H-SECTION", "I-SECTION",
                "W36X135", "W36X170", "W36", "W30",
                "Wide Flange", "UB", "IPE");

            // DAM DOC
            _symLong = FindSymbol(BuiltInCategory.OST_StructuralFraming,
                "THEP_H", "H-SECTION", "W12X26", "W12",
                "HW300", "IPE", "UB") ?? _symRafter;

            // XA GO MAI C200
            _symPurlin = FindSymbol(BuiltInCategory.OST_StructuralFraming,
                "THEP_C", "THEP_U", "C-CHANNEL", "C200",
                "C8X11P5", "C8", "Channel", "UPN") ?? _symLong;

            // XA GO TUONG C180
            _symGirt = FindSymbol(BuiltInCategory.OST_StructuralFraming,
                "THEP_C", "THEP_U", "C-CHANNEL", "C180",
                "C7X9P8", "C7", "UPN") ?? _symPurlin;

            // GIANG
            _symBrace = FindSymbol(BuiltInCategory.OST_StructuralFraming,
                "THEP_ONG", "THEP_HOP", "HSS", "PIPE",
                "HSS3X3", "HSS3", "RHS", "CHS") ?? _symLong;

            // CUA
            _symDoor = FindSymbol(BuiltInCategory.OST_Doors,
                "CUA_DON", "CUA_1CANH", "M_Single-Flush",
                "Single-Flush", "Door-Single", "Door");
            _symDoorDbl = FindSymbol(BuiltInCategory.OST_Doors,
                "CUA_DOI", "CUA_2CANH", "M_Double-Flush",
                "Double-Flush", "Door-Double") ?? _symDoor;
            _symDoorRoll = FindSymbol(BuiltInCategory.OST_Doors,
                "CUA_CUON", "CUA_NANG", "Roller", "Rolling",
                "Overhead", "SHUTTER") ?? _symDoorDbl;

            // CUA SO
            _symWindow = FindSymbol(BuiltInCategory.OST_Windows,
                "CUA_SO", "WINDOW", "CuaSo", "Window", "W-"
            ) ?? null;

            // Activate
            FamilySymbol[] allSym = {
                _symColBien, _symColGiua, _symColHoi, _symRafter,
                _symLong, _symPurlin, _symGirt, _symBrace,
                _symDoor, _symDoorDbl, _symDoorRoll, _symWindow };
            foreach (FamilySymbol fs in allSym)
                if (fs != null && !fs.IsActive) fs.Activate();

            // Log
            LogSym("ColBien", _symColBien);
            LogSym("ColGiua", _symColGiua);
            LogSym("ColHoi", _symColHoi);
            LogSym("Rafter", _symRafter);
            LogSym("DamDoc", _symLong);
            LogSym("Purlin", _symPurlin);
            LogSym("Girt", _symGirt);
            LogSym("Brace", _symBrace);
            LogSym("Door", _symDoor);
            LogSym("Window", _symWindow);
        }

        private void LogSym(string role, FamilySymbol fs)
        {
            if (fs != null)
                System.Diagnostics.Debug.WriteLine(
                    string.Format("[BIM] Symbol {0} -> {1}:{2}", role, fs.FamilyName, fs.Name));
            else
                Warn("Symbol " + role + ": NULL – khong tim thay trong template");
        }

        private FamilySymbol FindSymbol(BuiltInCategory cat, params string[] kws)
        {
            List<FamilySymbol> all = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol)).OfCategory(cat)
                .Cast<FamilySymbol>().ToList();
            foreach (string kw in kws)
            {
                FamilySymbol found = all.FirstOrDefault(fs =>
                    fs.FamilyName.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    fs.Name.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0);
                if (found != null) return found;
            }
            return all.FirstOrDefault();
        }

        // ======================================================================
        // 5. TYPES – Uu tien type VN tu template HUTECH
        // ======================================================================
        private void ResolveTypes()
        {
            _wtExt = FindType<WallType>(
                "Ton", "TON_SONG", "AZ100", "Metal_Panel",
                "TUONG_TON", "Metal", "Exterior", "Basic Wall");
            _wtInt = FindType<WallType>(
                "TUONG_THACH_CAO", "TUONG_NHE", "Interior", "Partition",
                "Gypsum", "Generic") ?? _wtExt;
            _wtBrick = FindType<WallType>(
                "TUONG_GACH", "Brick", "Gach", "GACH_220", "220",
                "Masonry") ?? _wtInt;
            _ftNen = FindType<FloorType>(
                "SAN_BTCT", "SAN_B25", "Concrete", "BETONG",
                "150", "Slab", "Generic Floor");
            _ftRanh = FindType<FloorType>(
                "RANH", "DRAIN", "FRP", "Ranh_Thoat") ?? _ftNen;
            _roofT = FindType<RoofType>(
                "MAI_TON", "TON_MAI", "Metal_Roof", "Metal",
                "Ton", "Basic", "Generic Roof");
        }

        private T FindType<T>(params string[] kws) where T : ElementType
        {
            List<T> all = new FilteredElementCollector(_doc)
                .OfClass(typeof(T)).Cast<T>().ToList();
            foreach (string kw in kws)
            {
                T found = all.FirstOrDefault(t =>
                    t.Name.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0);
                if (found != null) return found;
            }
            return all.FirstOrDefault();
        }

        // ======================================================================
        // 6. FOUNDATIONS
        // ======================================================================
        private void CreateFoundations()
        {
            foreach (double xM in K.GX)
            {
                foreach (double yM in K.GY)
                {
                    string pos = K.LX(xM) + K.LY(yM);
                    try
                    {
                        MkBox(K.pt(xM, yM, K.Z_MONG),
                            K.FD_L + 0.4, K.FD_W + 0.4, K.FD_LOT_T,
                            BuiltInCategory.OST_StructuralFoundation,
                            "MG_LOT_" + pos, "LOT_2400x2400x100_BTgachvung");

                        MkBox(K.pt(xM, yM, K.Z_MONG + K.FD_LOT_T),
                            K.FD_L, K.FD_W, K.FD_H,
                            BuiltInCategory.OST_StructuralFoundation,
                            "MG_DAI_" + pos, "DAI_2000x2000x500_B25");

                        double zCoBot = K.Z_MONG + K.FD_LOT_T + K.FD_H;
                        double hCo = K.Z_DINH_CO - zCoBot;
                        if (hCo > 0.01)
                        {
                            MkBox(K.pt(xM, yM, zCoBot),
                                K.FD_COL_W, K.FD_COL_H, hCo,
                                BuiltInCategory.OST_StructuralFoundation,
                                "MG_CO_" + pos, "CO_400x400x250_B25");
                        }
                        _nFnd++;
                    }
                    catch (Exception ex) { Warn("FND " + pos + ": " + ex.Message); }
                }
            }

            // Mong cot hoi CH
            foreach (double yM in new[] { 0.0, K.L_TOTAL })
            {
                for (int i = 0; i < K.GX.Length - 1; i++)
                {
                    double xMid = (K.GX[i] + K.GX[i + 1]) / 2.0;
                    MkBox(K.pt(xMid, yM, K.Z_MONG + K.FD_LOT_T),
                        1.2, 1.2, 0.35,
                        BuiltInCategory.OST_StructuralFoundation,
                        string.Format("MG_CH_X{0:F0}_{1}", xMid, K.LY(yM)),
                        "MONG_CH_1200x1200x350_B25");
                    _nFnd++;
                }
            }
        }

        // ======================================================================
        // 7. COLUMNS
        // ======================================================================
        private void CreateColumns()
        {
            if (_symColBien == null) { Warn("Col: thieu symbol"); return; }

            foreach (double xM in K.GX)
            {
                bool bien = Math.Abs(xM) < 0.01 || Math.Abs(xM - K.W_TOTAL) < 0.01;
                FamilySymbol sym = bien ? _symColBien : (_symColGiua ?? _symColBien);

                foreach (double yM in K.GY)
                {
                    try
                    {
                        FamilyInstance col = FamilyInstance.Create(
                            _doc, K.pt(xM, yM, K.Z_NEN), sym.Id, _lvNen.Id,
                            StructuralType.Column);
                        col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
                            ?.Set(_lvEave.Id);
                        col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
                            ?.Set(0.0);
                        col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
                            ?.Set(0.0);

                        string mk = string.Format("{0}_{1}{2}",
                            bien ? "KC_C1" : "KC_C2", K.LX(xM), K.LY(yM));
                        SetBimTag(col, mk, "COT_" + mk + "_NGAM_TCVN5575");

                        _cols.Add(col);
                        _nCol++;
                    }
                    catch (Exception ex)
                    {
                        Warn("Col " + K.LX(xM) + K.LY(yM) + ": " + ex.Message);
                    }
                }
            }

            // Cot hoi CH
            if (_symColHoi == null) return;
            foreach (double yM in new[] { 0.0, K.L_TOTAL })
            {
                for (int i = 0; i < K.GX.Length - 1; i++)
                {
                    double xMid = (K.GX[i] + K.GX[i + 1]) / 2.0;
                    try
                    {
                        FamilyInstance col = FamilyInstance.Create(
                            _doc, K.pt(xMid, yM, K.Z_NEN), _symColHoi.Id, _lvNen.Id,
                            StructuralType.Column);
                        col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
                            ?.Set(_lvEave.Id);
                        col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
                            ?.Set(0.0);
                        string mk = string.Format("KC_CH_X{0:F0}_{1}", xMid, K.LY(yM));
                        SetBimTag(col, mk, "COT_HOI_CH_H320x186_TCVN5575");
                        _cols.Add(col);
                        _nCol++;
                    }
                    catch (Exception ex) { Warn("ColHoi: " + ex.Message); }
                }
            }
        }

        // ======================================================================
        // 8. BEAMS
        // ======================================================================
        private void CreateBeams()
        {
            if (_symRafter == null) { Warn("Beam: thieu symbol"); return; }
            double[] nodeX = { 0, 8, 16, K.RIDGE_X, 24, 32, 40 };

            // Keo chinh
            foreach (double yM in K.GY)
            {
                for (int i = 0; i < nodeX.Length - 1; i++)
                {
                    double x1 = nodeX[i], x2 = nodeX[i + 1];
                    string mk = string.Format("KC_KEO_{0}_{1}", K.LY(yM), i + 1);
                    FamilyInstance fi = MakeBeam(
                        K.pt(x1, yM, K.RoofZ(x1)),
                        K.pt(x2, yM, K.RoofZ(x2)),
                        _symRafter, _lvEave, StructuralType.Beam,
                        mk, "KEO_K1.1_H(1020-570)x214", 0, 0);
                    if (fi != null) { _rafters.Add(fi); _nRaf++; }
                }
            }

            // Keo dau hoi
            foreach (double yM in new[] { 0.0, K.L_TOTAL })
            {
                for (int i = 0; i < K.GX.Length - 1; i++)
                {
                    double xM = (K.GX[i] + K.GX[i + 1]) / 2.0;
                    FamilySymbol rafSym = _symLong ?? _symRafter;
                    MakeBeam(
                        K.pt(K.GX[i], yM, K.RoofZ(K.GX[i])),
                        K.pt(xM, yM, K.RoofZ(xM)),
                        rafSym, _lvEave, StructuralType.Beam,
                        string.Format("KC_KEO_HOI_{0}_{1}a", K.LY(yM), K.GXN[i]),
                        "KEO_DAU_HOI", 1, 1);
                    MakeBeam(
                        K.pt(xM, yM, K.RoofZ(xM)),
                        K.pt(K.GX[i + 1], yM, K.RoofZ(K.GX[i + 1])),
                        rafSym, _lvEave, StructuralType.Beam,
                        string.Format("KC_KEO_HOI_{0}_{1}b", K.LY(yM), K.GXN[i]),
                        "KEO_DAU_HOI", 1, 1);
                    _nRaf += 2;
                }
            }

            // Dam doc + eave strut + ridge tie
            for (int i = 0; i < K.GY.Length - 1; i++)
            {
                double y1 = K.GY[i], y2 = K.GY[i + 1];
                foreach (double xM in K.GX)
                {
                    double zM = K.RoofZ(xM);
                    bool eave = Math.Abs(xM) < 0.01 || Math.Abs(xM - K.W_TOTAL) < 0.01;
                    string kind = eave ? "KC_EAVE" : "KC_DDC";
                    MakeBeam(K.pt(xM, y1, zM), K.pt(xM, y2, zM),
                        eave ? (_symPurlin ?? _symLong) : _symLong,
                        _lvEave, StructuralType.Beam,
                        string.Format("{0}_{1}_{2}_{3}", kind, K.LX(xM), K.LY(y1), K.LY(y2)),
                        kind, 0, 0);
                    _nRaf++;
                }
                MakeBeam(K.pt(K.RIDGE_X, y1, K.Z_RIDGE),
                    K.pt(K.RIDGE_X, y2, K.Z_RIDGE),
                    _symLong, _lvRidge, StructuralType.Beam,
                    string.Format("KC_RIDGE_{0}_{1}", K.LY(y1), K.LY(y2)),
                    "RIDGE_TIE", 0, 0);
                _nRaf++;
            }
        }

        // ======================================================================
        // 9. PURLINS
        // ======================================================================
        private void CreatePurlins()
        {
            if (_symPurlin == null) { Warn("Purlin: thieu symbol"); return; }
            List<double> pxList = CalcPurlinX();

            for (int i = 0; i < K.GY.Length - 1; i++)
            {
                double y1 = K.GY[i], y2 = K.GY[i + 1];
                foreach (double xM in pxList)
                {
                    FamilyInstance fi = MakeBeam(
                        K.pt(xM, y1, K.RoofZ(xM)),
                        K.pt(xM, y2, K.RoofZ(xM)),
                        _symPurlin, _lvEave, StructuralType.Beam,
                        string.Format("KCP_XGM_X{0:F2}_{1}_{2}", xM, K.LY(y1), K.LY(y2)),
                        "XA_GO_C200_1400", 1, 1);
                    if (fi != null) { _purlins.Add(fi); _nPur++; }
                }
            }

            // Xa go tuong
            if (_symGirt == null) return;
            foreach (double zM in K.GIRT_Z)
            {
                for (int i = 0; i < K.GY.Length - 1; i++)
                {
                    double y1 = K.GY[i], y2 = K.GY[i + 1];
                    foreach (double xM in new[] { 0.0, K.W_TOTAL })
                    {
                        MakeBeam(K.pt(xM, y1, zM), K.pt(xM, y2, zM),
                            _symGirt, _lvNen, StructuralType.Beam,
                            string.Format("KCP_XGT_{0}_Z{1:F1}_{2}_{3}",
                                K.LX(xM), zM, K.LY(y1), K.LY(y2)),
                            "XA_GO_TUONG_C180", 1, 1);
                        _nGirt++;
                    }
                }
                for (int i = 0; i < K.GX.Length - 1; i++)
                {
                    double x1 = K.GX[i], x2 = K.GX[i + 1];
                    foreach (double yM in new[] { 0.0, K.L_TOTAL })
                    {
                        MakeBeam(K.pt(x1, yM, zM), K.pt(x2, yM, zM),
                            _symGirt, _lvNen, StructuralType.Beam,
                            string.Format("KCP_XGT_HOI_{0}_Z{1:F1}_{2}_{3}",
                                K.LY(yM), zM, K.LX(x1), K.LX(x2)),
                            "XA_GO_TUONG_DAU_HOI_C180", 1, 1);
                        _nGirt++;
                    }
                }
            }
        }

        private List<double> CalcPurlinX()
        {
            List<double> result = new List<double>();
            double dx = K.PURLIN_SP / Math.Sqrt(1 + K.SLOPE * K.SLOPE);
            double[] pri = { 0, 8, 16, K.RIDGE_X, 24, 32, 40 };
            for (double x = dx; x < K.RIDGE_X - dx * 0.5; x += dx)
                if (!pri.Any(s => Math.Abs(s - x) <= 0.25))
                    result.Add(Math.Round(x, 3));
            for (double x = K.RIDGE_X + dx; x < K.W_TOTAL - dx * 0.5; x += dx)
                if (!pri.Any(s => Math.Abs(s - x) <= 0.25))
                    result.Add(Math.Round(x, 3));
            return result;
        }

        // ======================================================================
        // 10. SAG RODS
        // ======================================================================
        private void CreateSagRods()
        {
            if (_symBrace == null) return;
            List<double> pxList = CalcPurlinX();
            pxList.Sort();

            List<double> allX = new List<double> { 0 };
            allX.AddRange(pxList);
            if (!allX.Contains(K.RIDGE_X)) allX.Add(K.RIDGE_X);
            allX.Add(K.W_TOTAL);
            allX.Sort();
            allX = allX.Distinct().ToList();

            for (int bay = 0; bay < K.GY.Length - 1; bay++)
            {
                double y1 = K.GY[bay];
                double yLen = K.GY[bay + 1] - y1;
                foreach (double yM in new[] { y1 + yLen / 3.0, y1 + yLen * 2.0 / 3.0 })
                {
                    for (int j = 0; j < allX.Count - 1; j++)
                    {
                        double x1 = allX[j], x2 = allX[j + 1];
                        if (Math.Abs(x2 - x1) < 0.3) continue;
                        if (K.GX.Any(gx => Math.Abs(gx - x1) < 0.1) &&
                            K.GX.Any(gx => Math.Abs(gx - x2) < 0.1)) continue;

                        FamilyInstance sr = MakeBeam(
                            K.pt(x1, yM, K.RoofZ(x1)),
                            K.pt(x2, yM, K.RoofZ(x2)),
                            _symBrace, _lvEave, StructuralType.Brace,
                            string.Format("KCP_SAG_Bay{0}_Y{1:F1}_X{2:F1}", bay + 1, yM, x1),
                            "SAG_ROD_D12_TCVN5575", 1, 1);
                        if (sr != null) { _nSagRod++; _nConn++; }
                    }
                }
            }
        }

        // ======================================================================
        // 11. BRACING
        // ======================================================================
        private void CreateBracing()
        {
            if (_symBrace == null) { Warn("Brace: thieu symbol"); return; }

            foreach (int bay in K.BRACE_BAY)
            {
                if (bay < 0 || bay >= K.GY.Length - 1) continue;
                double y1 = K.GY[bay], y2 = K.GY[bay + 1];

                foreach (double xM in new[] { 0.0, K.W_TOTAL })
                {
                    double tz = K.RoofZ(xM);
                    MakeBeam(K.pt(xM, y1, K.Z_NEN), K.pt(xM, y2, tz),
                        _symBrace, _lvNen, StructuralType.Brace,
                        string.Format("KCP_GNG_V_{0}_B{1}_D1", K.LX(xM), bay + 1),
                        "GIANG_DUNG_X", 1, 1);
                    MakeBeam(K.pt(xM, y1, tz), K.pt(xM, y2, K.Z_NEN),
                        _symBrace, _lvNen, StructuralType.Brace,
                        string.Format("KCP_GNG_V_{0}_B{1}_D2", K.LX(xM), bay + 1),
                        "GIANG_DUNG_X", 1, 1);
                    _nBrace += 2;
                }

                MakeRoofBrace(y1, y2, 0.0, K.RIDGE_X, "KCP_GNG_R_L_B" + (bay + 1));
                MakeRoofBrace(y1, y2, K.W_TOTAL, K.RIDGE_X, "KCP_GNG_R_R_B" + (bay + 1));
            }

            // Giang dau hoi
            foreach (double yM in new[] { 0.0, K.L_TOTAL })
            {
                for (int i = 0; i < K.GX.Length - 1; i++)
                {
                    double x1 = K.GX[i], x2 = K.GX[i + 1];
                    double xMid = (x1 + x2) / 2.0;
                    if (i == 0 || i == K.GX.Length - 2)
                    {
                        MakeBeam(K.pt(x1, yM, K.RoofZ(x1)),
                            K.pt(xMid, yM, K.RoofZ(xMid)),
                            _symBrace, _lvEave, StructuralType.Brace,
                            string.Format("KCP_GNG_HOI_{0}_{1}_1", K.LY(yM), K.GXN[i]),
                            "GIANG_DAU_HOI", 1, 1);
                        MakeBeam(K.pt(xMid, yM, K.RoofZ(xMid)),
                            K.pt(x2, yM, K.RoofZ(x2)),
                            _symBrace, _lvEave, StructuralType.Brace,
                            string.Format("KCP_GNG_HOI_{0}_{1}_2", K.LY(yM), K.GXN[i]),
                            "GIANG_DAU_HOI", 1, 1);
                        _nBrace += 2;
                    }
                }
            }
        }

        private void MakeRoofBrace(double y1, double y2, double xE, double xR, string tag)
        {
            double zE = K.RoofZ(xE);
            MakeBeam(K.pt(xE, y1, zE), K.pt(xR, y2, K.Z_RIDGE),
                _symBrace, _lvEave, StructuralType.Brace, tag + "_1", "GIANG_MAI_X", 1, 1);
            MakeBeam(K.pt(xR, y1, K.Z_RIDGE), K.pt(xE, y2, zE),
                _symBrace, _lvEave, StructuralType.Brace, tag + "_2", "GIANG_MAI_X", 1, 1);
            _nBrace += 2;
        }

        // ======================================================================
        // 12. CONNECTIONS
        // ======================================================================
        private void CreateConnections()
        {
            ConnBase();
            ConnBeamToCol();
            ConnBeamToBeam();
            ConnRafterToCol();
            ConnPurlinClip();
            ConnRidgeSplice();
            ConnBraceGusset();
            ConnFlyBrace();
        }

        private void ConnBase()
        {
            foreach (double xM in K.GX)
            {
                foreach (double yM in K.GY)
                {
                    string pos = K.LX(xM) + K.LY(yM);
                    try
                    {
                        MkBox(K.pt(xM, yM, K.Z_DINH_CO),
                            K.BP_L + 0.02, K.BP_W + 0.02, K.GROUT_T,
                            _catConn,
                            "LK_GRT_" + pos, "GROUT_VUA_50mm");

                        MkBox(K.pt(xM, yM, K.Z_NEN),
                            K.BP_L, K.BP_W, K.BP_T,
                            _catConn,
                            "LK_BP_" + pos, "BP_650x550x30_Q345");

                        for (int s = -1; s <= 1; s += 2)
                        {
                            double off = K.COL_TW / 2.0 + K.BP_STF_T / 2.0 + 0.01;
                            MkBox(K.pt(xM + s * off, yM, K.Z_NEN + K.BP_T),
                                K.BP_STF_T, K.BP_W * 0.7, K.BP_STF_H,
                                _catConn,
                                "LK_STF_Y_" + pos + (s > 0 ? "_R" : "_L"), "SUON_GAN_Y_t16");
                            MkBox(K.pt(xM, yM + s * K.COL_B / 4.0, K.Z_NEN + K.BP_T),
                                K.BP_L * 0.7, K.BP_STF_T, K.BP_STF_H,
                                _catConn,
                                "LK_STF_X_" + pos + (s > 0 ? "_T" : "_B"), "SUON_GAN_X_t16");
                            _nConn += 2;
                        }

                        double hL = K.AB_PL / 2.0, hW = K.AB_PW / 2.0;
                        double[][] pat = {
                            new[]{ hL, hW}, new[]{ hL,-hW},
                            new[]{-hL, hW}, new[]{-hL,-hW} };
                        for (int b = 0; b < 4; b++)
                        {
                            double bx = xM + pat[b][0], by = yM + pat[b][1];
                            MkCyl(K.pt(bx, by, K.Z_NEN - K.AB_L),
                                K.AB_D / 2, K.AB_L,
                                _catConn,
                                string.Format("LK_ABL_{0}_{1}", pos, b + 1),
                                "BL_NEO_M30_L700_Q345");
                            MkBox(K.pt(bx, by, K.Z_NEN + K.BP_T),
                                K.AB_D * 2, K.AB_D * 2, K.AB_D,
                                _catConn,
                                string.Format("LK_NUT_{0}_{1}", pos, b + 1), "DAI_OC_M30");
                            _nConn += 2;
                        }
                    }
                    catch (Exception ex) { Warn("Base " + pos + ": " + ex.Message); }
                }
            }
        }

        private void ConnBeamToCol()
        {
            for (int i = 0; i < K.GY.Length - 1; i++)
            {
                double y1 = K.GY[i];
                foreach (double xM in K.GX)
                {
                    double zM = K.RoofZ(xM);
                    string pos = K.LX(xM) + "_" + K.LY(y1);
                    try
                    {
                        MkVPlate(K.pt(xM, y1, zM - K.EP_H / 2),
                            K.EP_H, K.EP_W, K.EP_T, false,
                            "LK_EP_B2C_" + pos, "EP_400x200x20_Q345");
                        _nConn++;
                        for (int r = 0; r < 3; r++)
                            for (int s = -1; s <= 1; s += 2)
                            {
                                MkCyl(K.pt(xM, y1 + s * 0.05, zM + 0.12 - r * 0.12),
                                    K.BOLT_M20 / 2, K.BOLT_L,
                                    _catConn,
                                    string.Format("LK_BL20_B2C_{0}_R{1}_{2}",
                                        pos, r, s > 0 ? "R" : "L"),
                                    "BL_M20_8.8");
                                _nConn++;
                            }
                    }
                    catch (Exception ex) { Warn("B2C " + pos + ": " + ex.Message); }
                }
            }
        }

        private void ConnBeamToBeam()
        {
            for (int i = 0; i < K.GY.Length - 1; i++)
            {
                double yM = (K.GY[i] + K.GY[i + 1]) / 2.0;
                foreach (double xM in K.GX)
                {
                    double zM = K.RoofZ(xM);
                    string pos = K.LX(xM) + "_" + (i + 1);
                    try
                    {
                        MkBox(K.pt(xM - K.FP_T / 2, yM, zM - K.FP_H / 2),
                            K.FP_T, K.FP_W, K.FP_H,
                            _catConn,
                            "LK_FP_" + pos, "FP_250x120x12_Q345");
                        _nConn++;
                        for (int b = 0; b < 3; b++)
                        {
                            MkCyl(K.pt(xM, yM, zM - 0.05 - b * 0.08),
                                K.BOLT_M16 / 2, K.BOLT_L * 0.8,
                                _catConn,
                                string.Format("LK_BL16_B2B_{0}_{1}", pos, b),
                                "BL_M16_8.8");
                            _nConn++;
                        }
                    }
                    catch (Exception ex) { Warn("B2B " + pos + ": " + ex.Message); }
                }
            }
        }

        private void ConnRafterToCol()
        {
            foreach (double xM in K.GX)
            {
                double ez = K.RoofZ(xM);
                bool isA = Math.Abs(xM) < 0.01;
                bool isF = Math.Abs(xM - K.W_TOTAL) < 0.01;
                List<double> dirs = new List<double>();
                if (isA) dirs.Add(1.0);
                else if (isF) dirs.Add(-1.0);
                else { dirs.Add(-1.0); dirs.Add(1.0); }

                foreach (double yM in K.GY)
                {
                    foreach (double dx in dirs)
                    {
                        string sd = dx > 0 ? "R" : "L";
                        string pos = K.LX(xM) + K.LY(yM) + "_" + sd;
                        try
                        {
                            MakeHaunch(xM, yM, ez, dx, pos);
                            MkBox(K.pt(xM + dx * K.GS_W * 0.4, yM, ez - K.GS_H * 0.5),
                                K.GS_W, K.GS_T, K.GS_H,
                                _catConn,
                                "LK_GS_" + pos, "BAN_MA_600x450x20_Q345");
                            _nConn++;
                            MkVPlate(K.pt(xM, yM, ez - K.EP_H * 0.6),
                                K.EP_H * 1.5, K.EP_W * 1.2, K.EP_T, false,
                                "LK_EPK_" + pos, "EP_KNEE_Q345");
                            _nConn++;
                            for (int r = 0; r < 5; r++)
                                for (int s = -1; s <= 1; s += 2)
                                {
                                    MkCyl(K.pt(xM, yM + s * 0.055,
                                            ez - 0.05 - r * 0.10),
                                        K.BOLT_M20 / 2, K.BOLT_L,
                                        _catConn,
                                        string.Format("LK_BLK_{0}_R{1}_{2}",
                                            pos, r, s > 0 ? "R" : "L"),
                                        "BL_M20_10.9");
                                    _nConn++;
                                }
                            MkBox(K.pt(xM, yM, ez - K.EP_H * 0.3),
                                K.COL_B, K.COL_TW + 0.008, K.EP_H,
                                _catConn,
                                "LK_WST_" + pos, "SUON_GAN_BUONG_COT_Q345");
                            _nConn++;
                        }
                        catch (Exception ex) { Warn("R2C " + pos + ": " + ex.Message); }
                    }
                }
            }
        }

        private void ConnPurlinClip()
        {
            List<double> pxList = CalcPurlinX();
            for (int i = 0; i < K.GY.Length - 1; i++)
            {
                double y1 = K.GY[i], y2 = K.GY[i + 1];
                foreach (double xM in pxList)
                {
                    double zM = K.RoofZ(xM);
                    foreach (double yC in new[] { y1, y2 })
                    {
                        string pos = string.Format("X{0:F1}_{1}", xM, K.LY(yC));
                        try
                        {
                            MkBox(K.pt(xM - K.CA_T / 2, yC, zM - K.CA_W / 2),
                                K.CA_T, K.CA_L, K.CA_W,
                                _catConn,
                                "LK_CA_" + pos, "THEP_GOC_L75x6_Q235");
                            _nConn++;
                            for (int b = 0; b < 2; b++)
                            {
                                MkCyl(K.pt(xM, yC + (b == 0 ? 0.04 : -0.04), zM),
                                    K.BOLT_M12 / 2, K.BOLT_L * 0.6,
                                    _catConn,
                                    string.Format("LK_BL12_{0}_{1}", pos, b),
                                    "BL_M12_4.8");
                                _nConn++;
                            }
                        }
                        catch (Exception ex) { Warn("CA " + pos + ": " + ex.Message); }
                    }
                }
            }
        }

        private void ConnRidgeSplice()
        {
            foreach (double yM in K.GY)
            {
                string pos = K.LY(yM);
                try
                {
                    MkBox(K.pt(K.RIDGE_X, yM, K.Z_RIDGE - K.RS_H / 2),
                        K.RS_T, K.RS_W, K.RS_H,
                        _catConn,
                        "LK_RS_" + pos, "BAN_NOI_NOC_350x300x20_Q345");
                    _nConn++;
                    for (int s = -1; s <= 1; s += 2)
                        for (int b = 0; b < 4; b++)
                        {
                            MkCyl(K.pt(K.RIDGE_X + s * 0.04, yM,
                                    K.Z_RIDGE - 0.04 - b * 0.07),
                                K.BOLT_M20 / 2, 0.06,
                                _catConn,
                                string.Format("LK_BLRS_{0}_{1}_{2}",
                                    pos, s > 0 ? "R" : "L", b),
                                "BL_M20_8.8");
                            _nConn++;
                        }
                }
                catch (Exception ex) { Warn("RS " + pos + ": " + ex.Message); }
            }
        }

        private void ConnBraceGusset()
        {
            foreach (int bay in K.BRACE_BAY)
            {
                if (bay < 0 || bay >= K.GY.Length - 1) continue;
                double y1 = K.GY[bay], y2 = K.GY[bay + 1];
                foreach (double xM in new[] { 0.0, K.W_TOTAL })
                {
                    double tz = K.RoofZ(xM);
                    foreach (double yM in new[] { y1, y2 })
                    {
                        MkBox(K.pt(xM, yM, K.Z_NEN), 0.30, 0.016, 0.30,
                            _catConn,
                            string.Format("LK_GS_BRC_{0}{1}", K.LX(xM), K.LY(yM)),
                            "BAN_MA_GIANG_CHAN");
                        MkBox(K.pt(xM, yM, tz - 0.15), 0.30, 0.016, 0.30,
                            _catConn,
                            string.Format("LK_GS_BRD_{0}{1}", K.LX(xM), K.LY(yM)),
                            "BAN_MA_GIANG_DINH");
                        _nConn += 2;
                    }
                }
            }
        }

        private void ConnFlyBrace()
        {
            List<double> pxList = CalcPurlinX();
            if (pxList.Count == 0) return;
            List<double> leftPx = pxList.Where(
                x => x > 0.5 && x < K.RIDGE_X).ToList();
            List<double> rightPx = pxList.Where(
                x => x > K.RIDGE_X && x < K.W_TOTAL - 0.5).ToList();
            if (leftPx.Count == 0 || rightPx.Count == 0) return;
            double firstLeft = leftPx.Min();
            double firstRight = rightPx.Min();

            foreach (double yM in K.GY)
            {
                foreach (double xPur in new[] { firstLeft, firstRight })
                {
                    double xCol = K.GX.OrderBy(gx => Math.Abs(gx - xPur)).First();
                    double zPur = K.RoofZ(xPur);
                    double zCol = K.RoofZ(xCol);
                    string nm = string.Format("FLY_BR_{0}_{1:F1}", K.LY(yM), xPur);
                    try
                    {
                        MkBox(K.pt((xPur + xCol) / 2, yM, (zPur + zCol) / 2 - 0.15),
                            Math.Abs(xPur - xCol) * 0.8, 0.010, 0.100,
                            _catConn,
                            nm, "KCP_FLY_BRACE_L75x6");
                        _nConn++;
                    }
                    catch { }
                }
            }
        }

        // ======================================================================
        // 13. FLOORS
        // ======================================================================
        private void CreateFloors()
        {
            if (_ftNen == null) { Warn("Floor: thieu type"); return; }
            MakeFloor(0, 0, K.W_TOTAL, K.L_TOTAL, _ftNen, _lvNen, "SN_NEN_EPOXY");
            MakeFloor(0, 0, 16, 17, _ftNen, _lvLung, "SN_LUNG_VP");

            if (_ftRanh == null) return;
            MakeFloor(15.85, 17, 16.15, 51, _ftRanh, _lvDinh, "SN_RANH_C_3-7");
            MakeFloor(16, 33.85, 32, 34.15, _ftRanh, _lvDinh, "SN_RANH_5_C-E");
            MakeFloor(16, 50.85, 40, 51.15, _ftRanh, _lvDinh, "SN_RANH_7_C-F");
            MakeFloor(31.85, 34, 32.15, 68, _ftRanh, _lvDinh, "SN_RANH_E_5-9");
            MakeFloor(-0.3, 0, 0, K.L_TOTAL, _ftRanh, _lvDinh, "SN_RANH_CV_A");
            MakeFloor(K.W_TOTAL, 0, K.W_TOTAL + 0.3, K.L_TOTAL, _ftRanh, _lvDinh, "SN_RANH_CV_F");
            MakeFloor(0, -0.3, K.W_TOTAL, 0, _ftRanh, _lvDinh, "SN_RANH_CV_1");
            MakeFloor(0, K.L_TOTAL, K.W_TOTAL, K.L_TOTAL + 0.3, _ftRanh, _lvDinh, "SN_RANH_CV_11");
        }

        private void MakeFloor(double x1, double y1, double x2, double y2,
                                FloorType ft, Level lv, string tag)
        {
            try
            {
                if (Math.Abs(x2 - x1) < 0.05 || Math.Abs(y2 - y1) < 0.05) return;
                List<CurveLoop> loops = new List<CurveLoop> { MakeRect(x1, y1, x2, y2) };
                Floor fl = Floor.Create(_doc, loops, ft.Id, lv.Id);
                SetBimTag(fl, tag, tag);
            }
            catch (Exception ex) { Warn("Floor " + tag + ": " + ex.Message); }
        }

        // ======================================================================
        // 14. WALLS
        // ======================================================================
        private void CreateWalls()
        {
            if (_wtExt == null) { Warn("Wall: thieu type"); return; }
            double hFt = K.ft(K.Z_EAVE);
            WallType wi = _wtInt ?? _wtExt;
            WallType wb = _wtBrick ?? wi;

            // Ngoai
            for (int i = 0; i < K.GY.Length - 1; i++)
            {
                MakeWall(0, K.GY[i], 0, K.GY[i + 1], hFt, _wtExt,
                    "TG_EXT_A_" + K.LY(K.GY[i]));
                MakeWall(K.W_TOTAL, K.GY[i], K.W_TOTAL, K.GY[i + 1], hFt, _wtExt,
                    "TG_EXT_F_" + K.LY(K.GY[i]));
            }
            for (int i = 0; i < K.GX.Length - 1; i++)
            {
                MakeWall(K.GX[i], 0, K.GX[i + 1], 0, hFt, _wtExt,
                    "TG_EXT_1_" + K.GXN[i]);
                MakeWall(K.GX[i], K.L_TOTAL, K.GX[i + 1], K.L_TOTAL, hFt, _wtExt,
                    "TG_EXT_11_" + K.GXN[i]);
            }

            // Noi that
            MakeWall(0, 8.5, 16, 8.5, hFt, wb, "TG_THAYDO_2_A-C");
            MakeWall(0, 17, 16, 17, hFt, wb, "TG_VP_3_A-C");
            MakeWall(8, 0, 8, 17, hFt, wb, "TG_INT_B_1-3");
            MakeWall(16, 0, 16, 17, hFt, wb, "TG_INT_C_1-3");
            MakeWall(16, 17, 16, 34, hFt, wi, "TG_INT_C_3-5");
            MakeWall(16, 25.5, 32, 25.5, hFt, wi, "TG_INT_4_C-E");
            MakeWall(32, 17, 32, 34, hFt, wi, "TG_INT_E_3-5");
            MakeWall(16, 34, 16, 51, hFt, wi, "TG_INT_C_5-7");
            MakeWall(24, 25.5, 24, 51, hFt, wi, "TG_INT_D_4-7");
            MakeWall(16, 42.5, 24, 42.5, hFt, wi, "TG_INT_6_C-D");
            MakeWall(24, 34, 32, 34, hFt, wi, "TG_INT_5_D-E");
            MakeWall(24, 51, 40, 51, hFt, wi, "TG_INT_7_D-F");
            MakeWall(32, 34, 32, 51, hFt, wi, "TG_INT_E_5-7");
            MakeWall(16, 51, 16, 68, hFt, wi, "TG_INT_C_7-9");
            MakeWall(16, 68, 40, 68, hFt, wi, "TG_INT_9_C-F");
            MakeWall(32, 51, 32, 68, hFt, wi, "TG_INT_E_7-9");
            MakeWall(16, 68, 16, 85, hFt, wi, "TG_INT_C_9-11");
            MakeWall(32, 68, 32, 85, hFt, wi, "TG_INT_E_9-11");
            MakeWall(32, 8.5, 40, 8.5, hFt, wi, "TG_INT_HL_2_E-F");
            MakeWall(32, 17, 40, 17, hFt, wi, "TG_INT_HL_3_E-F");
        }

        private void MakeWall(double x1, double y1, double x2, double y2,
                               double hFt, WallType wt, string tag)
        {
            try
            {
                XYZ p1 = new XYZ(K.ft(x1), K.ft(y1), K.ft(K.Z_NEN));
                XYZ p2 = new XYZ(K.ft(x2), K.ft(y2), K.ft(K.Z_NEN));
                if (p1.DistanceTo(p2) < K.ft(0.05)) return;
                Wall wall = Wall.Create(_doc, Line.CreateBound(p1, p2),
                    wt.Id, _lvNen.Id, hFt, 0, false, false);
                SetBimTag(wall, tag, tag);
                _nWall++;
            }
            catch (Exception ex) { Warn("Wall " + tag + ": " + ex.Message); }
        }

        // ======================================================================
        // 15. ROOF
        // [FIX-02] Doi ten bien 'cc' thanh 'ptC'
        // ======================================================================
        private void CreateRoof()
        {
            if (_roofT == null) { Warn("Roof: thieu type"); return; }
            try
            {
                double ez = K.ft(K.Z_EAVE);
                CurveArray footprint = new CurveArray();
                XYZ ptA = new XYZ(K.ft(0), K.ft(0), ez);
                XYZ ptB = new XYZ(K.ft(K.W_TOTAL), K.ft(0), ez);
                XYZ ptC = new XYZ(K.ft(K.W_TOTAL), K.ft(K.L_TOTAL), ez);
                XYZ ptD = new XYZ(K.ft(0), K.ft(K.L_TOTAL), ez);
                footprint.Append(Line.CreateBound(ptA, ptB));
                footprint.Append(Line.CreateBound(ptB, ptC));
                footprint.Append(Line.CreateBound(ptC, ptD));
                footprint.Append(Line.CreateBound(ptD, ptA));
                FootPrintRoof roof = FootPrintRoof.Create(
                    _doc, footprint, _lvEave, _roofT);
                if (roof == null) throw new InvalidOperationException("Roof null");
                // Architectural roof geometry created. In recent API the slope lines are assigned automatically.
                // Additional slope controls can be applied with roof.set_DefinesSlope/set_SlopeAngle if required.
                SetBimTag(roof, "MAI_TON_AZ100", "MAI_TON_AZ100_0.45mm_i20.7%_TCVN2737");
            }
            catch (Exception ex)
            {
                Warn("Roof footprint: " + ex.Message);
                try { CreateRoofExtrusion(); }
                catch (Exception ex2) { Warn("Roof extrusion: " + ex2.Message); }
            }
        }

        private void CreateRoofExtrusion()
        {
            CurveArray profile = new CurveArray();
            double ez = K.ft(K.Z_EAVE), rz = K.ft(K.Z_RIDGE);
            profile.Append(Line.CreateBound(
                new XYZ(K.ft(0), 0, ez),
                new XYZ(K.ft(K.RIDGE_X), 0, rz)));
            profile.Append(Line.CreateBound(
                new XYZ(K.ft(K.RIDGE_X), 0, rz),
                new XYZ(K.ft(K.W_TOTAL), 0, ez)));
            ReferencePlane rp = ReferencePlane.Create(
                _doc, new XYZ(0, 0, 0), new XYZ(1, 0, 0), new XYZ(0, 0, 1),
                _doc.ActiveView);
            ExtrusionRoof.Create(
                _doc, profile, rp, _roofT, K.ft(0), K.ft(K.L_TOTAL));
        }

        // ======================================================================
        // 16. DOORS
        // ======================================================================
        private void PlaceDoors()
        {
            if (_symDoor == null) { Warn("Door: thieu symbol"); return; }
            // Mat A
            PlDr(0, 4.25, "D1"); PlDr(0, 12.75, "D1"); PlDr(0, 21, "D3");
            PlDr(0, 30, "D1"); PlDr(0, 55, "D1"); PlDr(0, 63, "D1");
            PlDr(0, 72, "D1"); PlDr(0, 80, "D1");
            // Mat F
            PlDr(K.W_TOTAL, 4.25, "D1"); PlDr(K.W_TOTAL, 12.75, "D1");
            PlDr(K.W_TOTAL, 21, "D1"); PlDr(K.W_TOTAL, 30, "D1");
            PlDr(K.W_TOTAL, 38, "D1"); PlDr(K.W_TOTAL, 46, "D1");
            PlDr(K.W_TOTAL, 55, "D1"); PlDr(K.W_TOTAL, 63, "D1");
            PlDr(K.W_TOTAL, 72, "D3"); PlDr(K.W_TOTAL, 80, "D1");
            // Dau hoi 1
            PlDr(4, 0, "D2"); PlDr(12, 0, "D2"); PlDr(28, 0, "D2"); PlDr(36, 0, "D2");
            // Dau hoi 11
            PlDr(4, K.L_TOTAL, "D2"); PlDr(12, K.L_TOTAL, "D2");
            PlDr(28, K.L_TOTAL, "D2"); PlDr(36, K.L_TOTAL, "D2");
            // Noi that truc C
            PlDr(16, 21, "D3"); PlDr(16, 30, "D4"); PlDr(16, 38, "D4");
            PlDr(16, 46, "D4"); PlDr(16, 55, "D3"); PlDr(16, 63, "D3");
            PlDr(16, 72, "D3"); PlDr(16, 80, "D3");
            // Truc D
            PlDr(24, 30, "D4"); PlDr(24, 38, "D4"); PlDr(24, 46, "D6");
            // Truc E
            PlDr(32, 21, "D1"); PlDr(32, 30, "D1"); PlDr(32, 46, "D1");
            PlDr(32, 55, "D1"); PlDr(32, 72, "D3");
            // Ngang
            PlDr(4, 17, "D5"); PlDr(12, 17, "D5");
            PlDr(20, 34, "D3"); PlDr(28, 34, "D3");
            PlDr(20, 51, "D3"); PlDr(28, 51, "D3"); PlDr(36, 51, "D3");
            PlDr(20, 68, "D3"); PlDr(28, 68, "D3"); PlDr(36, 68, "D3");
        }

        private void PlDr(double xM, double yM, string typ)
        {
            try
            {
                Wall hostWall = FindNearWall(xM, yM);
                if (hostWall == null)
                {
                    Warn("Door " + typ + " X" + xM + " Y" + yM + ": no wall");
                    return;
                }
                FamilySymbol sym = _symDoor;
                if (typ == "D2" || typ == "D6")
                    sym = _symDoorRoll ?? _symDoorDbl ?? _symDoor;
                else if (typ == "D3" || typ == "D5")
                    sym = _symDoorDbl ?? _symDoor;
                if (sym != null && !sym.IsActive) sym.Activate();
                FamilyInstance fi = FamilyInstance.Create(
                    _doc, K.pt(xM, yM, K.Z_NEN), sym.Id, hostWall, _lvNen,
                    StructuralType.NonStructural);
                string mk = string.Format("CUA_{0}_X{1}_Y{2}", typ, xM, yM);
                SetBimTag(fi, mk, "CUA_" + typ + "_TCVN");
                _nDoor++;
            }
            catch (Exception ex) { Warn("Door " + typ + ": " + ex.Message); }
        }

        private Wall FindNearWall(double xM, double yM)
        {
            XYZ target = K.pt(xM, yM, K.Z_NEN);
            double minDist = double.MaxValue;
            Wall best = null;
            foreach (Element el in new FilteredElementCollector(_doc)
                .OfClass(typeof(Wall)).WhereElementIsNotElementType())
            {
                Wall w = el as Wall;
                if (w == null) continue;
                LocationCurve lc = w.Location as LocationCurve;
                if (lc == null) continue;
                IntersectionResult ir = lc.Curve.Project(target);
                if (ir == null) continue;
                double dist = ir.XYZPoint.DistanceTo(target);
                if (dist < minDist) { minDist = dist; best = w; }
            }
            return minDist < K.ft(1.5) ? best : null;
        }

        // ======================================================================
        // 17. MEP
        // ======================================================================
        private void CreateMEP()
        {
            double zDen = K.Z_EAVE - 0.5;

            for (double x = 2; x < K.W_TOTAL; x += 4)
                for (double y = 19; y < 66; y += 4)
                {
                    MkBox(K.pt(x, y, zDen), K.LIGHT_SZ, K.LIGHT_SZ, 0.05,
                        BuiltInCategory.OST_LightingFixtures,
                        string.Format("MEP_DEN_SX_{0:F0}_{1:F0}", x, y),
                        "DEN_LED_600x600_SX");
                    _nMep++;
                }

            for (double x = 2; x < 15; x += 3)
                for (double y = 2; y < 16; y += 3)
                {
                    MkBox(K.pt(x, y, 2.8), K.LIGHT_SZ, K.LIGHT_SZ, 0.05,
                        BuiltInCategory.OST_LightingFixtures,
                        string.Format("MEP_DEN_VP_{0:F0}_{1:F0}", x, y),
                        "DEN_LED_600x600_VP");
                    _nMep++;
                }

            if (_symBrace != null)
            {
                double zP = K.Z_DINH - 0.1;
                MakeBeam(K.pt(16, 17, zP), K.pt(16, 51, zP),
                    _symBrace, _lvDinh, StructuralType.Beam,
                    "MEP_ONG_C_3-7", "ONG_THOAT_PVC_D110", 1, 1);
                MakeBeam(K.pt(32, 34, zP), K.pt(32, 68, zP),
                    _symBrace, _lvDinh, StructuralType.Beam,
                    "MEP_ONG_E_5-9", "ONG_THOAT_PVC_D110", 1, 1);
                MakeBeam(K.pt(16, 34, zP), K.pt(32, 34, zP),
                    _symBrace, _lvDinh, StructuralType.Beam,
                    "MEP_ONG_5_C-E", "ONG_THOAT_PVC_D110", 1, 1);
                _nMep += 3;
            }

            MkBox(K.pt(0.15, 8.5, 1.2), 0.3, 0.8, 1.2,
                BuiltInCategory.OST_ElectricalEquipment,
                "MEP_TDT", "TU_DIEN_CHINH_3P_400A");
            MkBox(K.pt(16.15, 25.5, 1.2), 0.3, 0.6, 1.0,
                BuiltInCategory.OST_ElectricalEquipment,
                "MEP_TDPP", "TU_PHAN_PHOI_SX");
            _nMep += 2;

            MkBox(K.pt(20, 42.5, K.Z_NEN), 2.0, 1.5, 2.2,
                BuiltInCategory.OST_MechanicalEquipment,
                "MEP_ML_CN", "MAY_LANH_CN_20HP");
            MkBox(K.pt(10, 42.5, K.Z_EAVE), 0.8, 0.8, 0.6,
                BuiltInCategory.OST_MechanicalEquipment,
                "MEP_QH_1", "QUAT_HUT_CN_D800");
            MkBox(K.pt(30, 42.5, K.Z_EAVE), 0.8, 0.8, 0.6,
                BuiltInCategory.OST_MechanicalEquipment,
                "MEP_QH_2", "QUAT_HUT_CN_D800");
            _nMep += 3;
        }

        // ======================================================================
        // 17.5 STAIRS (Missing elements)
        // ======================================================================
        private void CreateStairs()
        {
            try
            {
                // Model stairs as a direct shape stair block with TCVN naming
                double stepCount = 16;
                double stepHeight = 0.18;
                double stepDepth = 0.28;
                double stepWidth = 1.20;
                double startX = 2.0;
                double startY = 2.0;
                for (int i = 0; i < stepCount; i++)
                {
                    double z = K.Z_NEN + i * stepHeight;
                    double depth = stepDepth * (i + 1);
                    MkBox(K.pt(startX + depth / 2.0, startY, z + stepHeight / 2.0),
                        depth, stepWidth, stepHeight,
                        BuiltInCategory.OST_GenericModel,
                        string.Format("MG_STAIR_STEP_{0:00}", i + 1),
                        "CAU_THANG_TCVN_" + (i + 1));
                    _nMep++; // count as MEP/aux element for reporting
                }
            }
            catch (Exception ex) { Warn("Stairs: " + ex.Message); }
        }

        // ======================================================================
        // 17.6 WINDOWS / OPENINGS
        // ======================================================================
        private void PlaceWindows()
        {
            if (_symWindow == null)
            {
                Warn("Windows: thieu symbol");
                return;
            }

            // place window openings on exterior walls
            List<Wall> extWalls = new FilteredElementCollector(_doc)
                .OfClass(typeof(Wall)).OfCategory(BuiltInCategory.OST_Walls)
                .Cast<Wall>().Where(w => w.LevelId == _lvNen.Id).ToList();

            foreach (Wall wall in extWalls)
            {
                LocationCurve lc = wall.Location as LocationCurve;
                if (lc == null || lc.Curve == null) continue;
                XYZ p0 = lc.Curve.GetEndPoint(0);
                XYZ p1 = lc.Curve.GetEndPoint(1);
                double len = p0.DistanceTo(p1);
                if (len < K.ft(1.0)) continue;
                // place one window in the mid span of exterior wall
                XYZ mid = (p0 + p1) / 2.0;
                try
                {
                    FamilyInstance win = FamilyInstance.Create(
                        _doc, mid, _symWindow.Id, wall, _lvNen, StructuralType.NonStructural);
                    SetBimTag(win, "TG_WINDOW_" + wall.Id.IntegerValue, "CUA_SO_TCVN");
                    _nDoor++;
                }
                catch (Exception ex) { Warn("Window place: " + ex.Message); }
            }
        }

        // ======================================================================
        // 18. GUTTER
        // ======================================================================
        private void CreateGutter()
        {
            double zG = K.Z_EAVE - K.GUTTER_H;

            foreach (double xM in new[] { 0.0, K.W_TOTAL })
            {
                for (int i = 0; i < K.GY.Length - 1; i++)
                {
                    double y1 = K.GY[i], y2 = K.GY[i + 1];
                    MkBox(K.pt(xM, (y1 + y2) / 2, zG),
                        K.GUTTER_W, y2 - y1, K.GUTTER_H,
                        BuiltInCategory.OST_GenericModel,
                        string.Format("MAI_MNG_{0}_{1}_{2}",
                            K.LX(xM), K.LY(y1), K.LY(y2)),
                        "MANG_NUOC_300x150");
                    _nMep++;
                }
            }

            foreach (double xM in new[] { 0.0, K.W_TOTAL })
            {
                foreach (double yM in K.GY)
                {
                    MkCyl(K.pt(xM, yM, K.Z_NEN),
                        K.DOWNPIPE_D / 2, K.Z_EAVE - K.Z_NEN,
                        BuiltInCategory.OST_GenericModel,
                        string.Format("MAI_OXI_{0}_{1}", K.LX(xM), K.LY(yM)),
                        "ONG_XOI_D114");
                    _nMep++;
                }
            }
        }

        // ======================================================================
        // 19. BIM VIEWS [V5.0-03]
        // ======================================================================
        private void CreateBimViews()
        {
            try
            {
                ViewFamilyType vft3d = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(v => v.ViewFamily == ViewFamily.ThreeDimensional);

                if (vft3d != null)
                {
                    try
                    {
                        View3D v3d = View3D.CreateIsometric(_doc, vft3d.Id);
                        if (v3d != null)
                        {
                            try { v3d.Name = "BIM_3D_TONG_THE"; }
                            catch { try { v3d.Name = "BIM_3D_TONG_THE_V5"; } catch { } }
                            v3d.DetailLevel = ViewDetailLevel.Fine;
                            _nView++;
                        }
                    }
                    catch (Exception ex) { Warn("View 3D: " + ex.Message); }
                }

                ViewFamilyType vftPlan = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(v => v.ViewFamily == ViewFamily.FloorPlan);

                ViewFamilyType vftStr = new FilteredElementCollector(_doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(v => v.ViewFamily == ViewFamily.StructuralPlan);

                ViewFamilyType vftUse = vftStr ?? vftPlan;
                if (vftUse != null)
                {
                    if (_lvNen != null)
                        TryCreatePlan(vftUse, _lvNen, "BIM_MB_NEN_+0.000");
                    if (_lvEave != null)
                        TryCreatePlan(vftUse, _lvEave, "BIM_MB_EAVE_+7.000");
                    if (_lvMong != null)
                        TryCreatePlan(vftUse, _lvMong, "BIM_MB_MONG_-0.900");
                }
            }
            catch (Exception ex) { Warn("BimViews: " + ex.Message); }
        }

        private void TryCreatePlan(ViewFamilyType vft, Level lv, string name)
        {
            try
            {
                ViewPlan vp = ViewPlan.Create(_doc, vft.Id, lv.Id);
                if (vp != null)
                {
                    try { vp.Name = name; }
                    catch { try { vp.Name = name + "_V5"; } catch { } }
                    vp.DetailLevel = ViewDetailLevel.Fine;
                    _nView++;
                }
            }
            catch (Exception ex) { Warn("Plan " + name + ": " + ex.Message); }
        }

        // ======================================================================
        // 20. BIM SCHEDULES [V5.0-04]
        // ======================================================================
        private void CreateBimSchedules()
        {
            TryCreateSchedule(
                BuiltInCategory.OST_StructuralColumns,
                "BIM_BANG_COT_THEP",
                new BuiltInParameter[] {
                    BuiltInParameter.ALL_MODEL_MARK,
                    BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
                    BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
                    BuiltInParameter.FAMILY_TOP_LEVEL_PARAM,
                    BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS
                });

            TryCreateSchedule(
                BuiltInCategory.OST_StructuralFraming,
                "BIM_BANG_DAM_KEO",
                new BuiltInParameter[] {
                    BuiltInParameter.ALL_MODEL_MARK,
                    BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
                    BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                    BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS
                });

            // Lien ket la DirectShape GenericModel voi mark bat dau LK_
            TryCreateSchedule(
                BuiltInCategory.OST_GenericModel,
                "BIM_BANG_LIEN_KET",
                new BuiltInParameter[] {
                    BuiltInParameter.ALL_MODEL_MARK,
                    BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS
                });
        }

        private void TryCreateSchedule(BuiltInCategory cat, string name,
                                         BuiltInParameter[] paramIds)
        {
            try
            {
                ViewSchedule vs = ViewSchedule.CreateSchedule(
                    _doc, new ElementId(cat));
                if (vs == null) return;
                try { vs.Name = name; }
                catch { try { vs.Name = name + "_V5"; } catch { } }

                ScheduleDefinition def = vs.Definition;
                IList<SchedulableField> allFields = def.GetSchedulableFields();

                foreach (BuiltInParameter bip in paramIds)
                {
                    ElementId targetId = new ElementId(bip);
                    foreach (SchedulableField sf in allFields)
                    {
                        if (sf.ParameterId == targetId)
                        {
                            try { def.AddField(sf); }
                            catch { }
                            break;
                        }
                    }
                }
                _nSched++;
            }
            catch (Exception ex) { Warn("Schedule " + name + ": " + ex.Message); }
        }

        // ======================================================================
        // 21. CLEANUP
        // ======================================================================
        private void CleanupModel()
        {
            try
            {
                List<Element> cols = new FilteredElementCollector(_doc)
                    .OfCategory(BuiltInCategory.OST_StructuralColumns)
                    .WhereElementIsNotElementType().ToList();
                List<Element> floors = new FilteredElementCollector(_doc)
                    .OfCategory(BuiltInCategory.OST_Floors)
                    .WhereElementIsNotElementType().ToList();
                List<Element> walls = new FilteredElementCollector(_doc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType().ToList();

                foreach (Element col in cols)
                {
                    foreach (Element fl in floors)
                    {
                        try
                        {
                            if (!JoinGeometryUtils.AreElementsJoined(_doc, col, fl))
                            { JoinGeometryUtils.JoinGeometry(_doc, col, fl); _nClean++; }
                        }
                        catch { }
                    }
                    foreach (Element wl in walls)
                    {
                        try
                        {
                            if (!JoinGeometryUtils.AreElementsJoined(_doc, col, wl))
                            { JoinGeometryUtils.JoinGeometry(_doc, col, wl); _nClean++; }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex) { Warn("Cleanup: " + ex.Message); }
        }

        // ======================================================================
        // 22. VALIDATE
        // ======================================================================
        private void ValidateModel()
        {
            List<Level> lvs = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level)).Cast<Level>().ToList();
            int lvOk = 0;
            foreach (Level lv in lvs)
            {
                double elev = lv.Elevation / K.M2F;
                foreach (double pe in K.PROJ_ELEVS)
                    if (Math.Abs(elev - pe) < 0.01) { lvOk++; break; }
            }
            if (lvOk < K.PROJ_ELEVS.Length)
                Warn(string.Format("Validate: {0}/{1} level khop",
                    lvOk, K.PROJ_ELEVS.Length));

            List<Grid> grds = new FilteredElementCollector(_doc)
                .OfClass(typeof(Grid)).Cast<Grid>().ToList();
            HashSet<string> need = new HashSet<string>();
            foreach (string nm in K.GXN) need.Add(nm);
            foreach (string nm in K.GYN) need.Add(nm);
            int grOk = 0;
            foreach (Grid grd in grds)
                if (need.Contains(grd.Name)) grOk++;
            if (grOk < need.Count)
                Warn(string.Format("Validate: {0}/{1} grid khop", grOk, need.Count));

            int nSC = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .WhereElementIsNotElementType().GetElementCount();
            if (nSC < 60)
                Warn(string.Format("Validate: {0} cot (can >= 66+10)", nSC));

            int nSF = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .WhereElementIsNotElementType().GetElementCount();
            if (nSF < 50)
                Warn(string.Format("Validate: {0} dam/keo (it)", nSF));

            // Dem lien ket bang mark prefix LK_ (DirectShape GenericModel)
            int nLK = 0;
            int nGM = 0;
            try
            {
                foreach (Element el in new FilteredElementCollector(_doc)
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .WhereElementIsNotElementType())
                {
                    nGM++;
                    Parameter pm = el.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
                    if (pm != null)
                    {
                        string mk = pm.AsString();
                        if (!string.IsNullOrEmpty(mk) && mk.StartsWith("LK_"))
                            nLK++;
                    }
                }
            }
            catch { }
            int nGM_other = nGM - nLK;
            if (nGM_other > 100)
                Warn(string.Format("Validate: {0} GenericModel khong LK_ (nen giam)", nGM_other));

            // Kiem tra naming BIM
            int nNoMark = 0;
            List<BuiltInCategory> checkCats = new List<BuiltInCategory> {
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Doors
            };
            foreach (BuiltInCategory chkCat in checkCats)
            {
                try
                {
                    foreach (Element el in new FilteredElementCollector(_doc)
                        .OfCategory(chkCat).WhereElementIsNotElementType())
                    {
                        Parameter pm = el.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
                        if (pm == null || string.IsNullOrEmpty(pm.AsString()))
                            nNoMark++;
                    }
                }
                catch { }
            }
            if (nNoMark > 0)
                Warn(string.Format("Validate: {0} cau kien thieu Mark", nNoMark));

            System.Diagnostics.Debug.WriteLine(string.Format(
                "[BIM] Validate: Lv={0} Gr={1} Col={2} Frame={3} Conn={4} GM={5}",
                lvs.Count, grds.Count, nSC, nSF, nLK, nGM));
            _nClean++;
        }

        // ======================================================================
        // HELPER METHODS
        // ======================================================================

        // [V4.1-06] BIM Tag: Mark + Comments
        private void SetBimTag(Element el, string mark, string comment)
        {
            if (el == null) return;
            string mk = mark.Length > 250 ? mark.Substring(0, 250) : mark;
            string cm = comment.Length > 250 ? comment.Substring(0, 250) : comment;
            try { el.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)?.Set(mk); }
            catch { }
            try { el.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.Set(cm); }
            catch { }
        }

        private void SetDsTag(DirectShape ds, string mark, string comment)
        {
            if (ds == null) return;
            string nm = mark.Length > 120 ? mark.Substring(0, 120) : mark;
            ds.SetName(nm);
            string cm = comment.Length > 250 ? comment.Substring(0, 250) : comment;
            try { ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.Set(cm); }
            catch { }
        }

        private FamilyInstance MakeBeam(XYZ p1, XYZ p2, FamilySymbol sym,
            Level lv, StructuralType st,
            string mark, string comment, int rS, int rE)
        {
            if (sym == null) return null;
            if (p1.DistanceTo(p2) < K.ft(0.05))
            { Warn("Beam too short: " + mark); return null; }
            try
            {
                FamilyInstance fi = FamilyInstance.Create(
                    _doc, Line.CreateBound(p1, p2), sym.Id, lv.Id, st);
                fi.get_Parameter(BuiltInParameter.STRUCTURAL_START_RELEASE_TYPE)?.Set(rS);
                fi.get_Parameter(BuiltInParameter.STRUCTURAL_END_RELEASE_TYPE)?.Set(rE);
                SetBimTag(fi, mark, comment);
                return fi;
            }
            catch (Exception ex)
            { Warn("Beam " + mark + ": " + ex.Message); return null; }
        }

        private void MkBox(XYZ ctr, double lM, double wM, double hM,
                            BuiltInCategory cat, string mark, string comment)
        {
            try
            {
                double halfL = K.ft(lM) / 2, halfW = K.ft(wM) / 2;
                XYZ p1 = new XYZ(ctr.X - halfL, ctr.Y - halfW, ctr.Z);
                XYZ p2 = new XYZ(ctr.X + halfL, ctr.Y - halfW, ctr.Z);
                XYZ p3 = new XYZ(ctr.X + halfL, ctr.Y + halfW, ctr.Z);
                XYZ p4 = new XYZ(ctr.X - halfL, ctr.Y + halfW, ctr.Z);
                CurveLoop lp = new CurveLoop();
                lp.Append(Line.CreateBound(p1, p2));
                lp.Append(Line.CreateBound(p2, p3));
                lp.Append(Line.CreateBound(p3, p4));
                lp.Append(Line.CreateBound(p4, p1));
                Solid solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                    new List<CurveLoop> { lp }, XYZ.BasisZ, K.ft(hM));
                DirectShape ds = DirectShape.CreateElement(
                    _doc, new ElementId(cat));
                ds.SetShape(new List<GeometryObject> { solid });
                SetDsTag(ds, mark, comment);
            }
            catch (Exception ex) { Warn("Box " + mark + ": " + ex.Message); }
        }

        private void MkVPlate(XYZ bc, double hM, double wM, double tM,
                               bool axisX, string mark, string comment)
        {
            XYZ center = new XYZ(
                bc.X + (axisX ? K.ft(wM) / 2 : K.ft(tM) / 2),
                bc.Y + (axisX ? K.ft(tM) / 2 : K.ft(wM) / 2),
                bc.Z + K.ft(hM) / 2);
            MkBox(center, axisX ? wM : tM, axisX ? tM : wM, hM,
                _catConn, mark, comment);
        }

        private void MkCyl(XYZ bc, double rM, double hM,
                            BuiltInCategory cat, string mark, string comment)
        {
            try
            {
                int nSeg = 8;
                double rad = K.ft(rM), height = K.ft(hM);
                if (rad < 1e-6 || height < 1e-6) return;

                List<XYZ> pts = new List<XYZ>();
                for (int j = 0; j < nSeg; j++)
                {
                    double ang = 2.0 * Math.PI * j / nSeg;
                    pts.Add(new XYZ(bc.X + rad * Math.Cos(ang),
                                    bc.Y + rad * Math.Sin(ang), bc.Z));
                }
                CurveLoop lp = new CurveLoop();
                for (int j = 0; j < nSeg; j++)
                {
                    XYZ pA = pts[j], pB = pts[(j + 1) % nSeg];
                    if (pA.DistanceTo(pB) < 1e-9) continue;
                    lp.Append(Line.CreateBound(pA, pB));
                }
                Solid solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                    new List<CurveLoop> { lp }, XYZ.BasisZ, height);
                DirectShape ds = DirectShape.CreateElement(
                    _doc, new ElementId(cat));
                ds.SetShape(new List<GeometryObject> { solid });
                SetDsTag(ds, mark, comment);
            }
            catch (Exception ex) { Warn("Cyl " + mark + ": " + ex.Message); }
        }

        private void MakeHaunch(double xM, double yM, double ez,
                                 double dx, string tag)
        {
            try
            {
                double len = K.HAUNCH_L, dep = K.HAUNCH_D, thk = K.BP_T;
                XYZ pA = K.pt(xM, yM - thk / 2, ez);
                XYZ pB = K.pt(xM + dx * len, yM - thk / 2, ez - dep);
                XYZ pC2 = K.pt(xM, yM - thk / 2, ez - dep);
                CurveLoop face = new CurveLoop();
                if (pA.DistanceTo(pB) > 1e-6) face.Append(Line.CreateBound(pA, pB));
                if (pB.DistanceTo(pC2) > 1e-6) face.Append(Line.CreateBound(pB, pC2));
                if (pC2.DistanceTo(pA) > 1e-6) face.Append(Line.CreateBound(pC2, pA));
                if (face.NumberOfCurves() < 3) return;
                Solid solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                    new List<CurveLoop> { face }, XYZ.BasisY, K.ft(thk));
                DirectShape ds = DirectShape.CreateElement(
                    _doc, new ElementId(_catConn));
                ds.SetShape(new List<GeometryObject> { solid });
                SetDsTag(ds, "LK_HNC_" + tag, "HAUNCH_KNEE_L1500_D500_Q345");
                _nConn++;
            }
            catch (Exception ex) { Warn("Haunch " + tag + ": " + ex.Message); }
        }

        private CurveLoop MakeRect(double x1, double y1, double x2, double y2)
        {
            double lx = Math.Min(x1, x2), ly = Math.Min(y1, y2);
            double ux = Math.Max(x1, x2), uy = Math.Max(y1, y2);
            XYZ a = new XYZ(K.ft(lx), K.ft(ly), 0);
            XYZ b = new XYZ(K.ft(ux), K.ft(ly), 0);
            XYZ c = new XYZ(K.ft(ux), K.ft(uy), 0);
            XYZ d = new XYZ(K.ft(lx), K.ft(uy), 0);
            CurveLoop lp = new CurveLoop();
            lp.Append(Line.CreateBound(a, b));
            lp.Append(Line.CreateBound(b, c));
            lp.Append(Line.CreateBound(c, d));
            lp.Append(Line.CreateBound(d, a));
            return lp;
        }
    }
}
// ==============================================================================
// HUONG DAN SU DUNG
// ==============================================================================
// 1. Revit > File > New > Project > Browse
//    > chon "R2024_TEMPLATE KET CAU HUTECH_THUC HANH.rte" > OK
// 2. Visual Studio > Class Library (.NET Framework 4.8) > BIMxuong2
//    > Add Reference: RevitAPI.dll, RevitAPIUI.dll
//    > Paste code nay vao Command.cs > Build
// 3. Copy BIMxuong2.addin vao %AppData%\Autodesk\Revit\Addins\2024\
// 4. Revit > Add-Ins > External Tools > BIM Xuong So 2 V4.1
//
// ==============================================================================
// .addin
// ==============================================================================
/*
<?xml version="1.0" encoding="utf-8"?>
<RevitAddIns>
  <AddIn Type="Command">
    <Name>BIM Xuong So 2 V4.1</Name>
    <Assembly>BIMxuong2.dll</Assembly>
    <AddInId>A1B2C3D4-E5F6-7890-ABCD-EF1234567890</AddInId>
    <FullClassName>BIMxuong2.Command</FullClassName>
    <VendorId>CUONGTHUY</VendorId>
    <VendorDescription>BIM Xuong Chan Ga Cay So 2 – TCVN – HUTECH</VendorDescription>
  </AddIn>
</RevitAddIns>
*/
// ==============================================================================
// .csproj
// ==============================================================================
/*
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Library</OutputType>
    <TargetFramework>net48</TargetFramework>
    <LangVersion>7.3</LangVersion>
    <Nullable>disable</Nullable>
    <AssemblyName>BIMxuong2</AssemblyName>
    <RootNamespace>BIMxuong2</RootNamespace>
    <PlatformTarget>x64</PlatformTarget>
  </PropertyGroup>
  <ItemGroup>
    <Reference Include="RevitAPI">
      <HintPath>C:\Program Files\Autodesk\Revit 2024\RevitAPI.dll</HintPath>
      <Private>False</Private>
    </Reference>
    <Reference Include="RevitAPIUI">
      <HintPath>C:\Program Files\Autodesk\Revit 2024\RevitAPIUI.dll</HintPath>
      <Private>False</Private>
    </Reference>
  </ItemGroup>
</Project>
*/