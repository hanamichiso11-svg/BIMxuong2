using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace BIMxuong2
{
    internal class XuongBuilder
    {
        private readonly Document _doc;

        private Level _lvBase;
        private Level _lvEave;
        private Level _lvRidge;

        private FamilySymbol _columnSymbol;
        private FamilySymbol _beamSymbol;
        private FamilySymbol _purlinSymbol;
        private WallType _wallType;
        private RoofType _roofType;

        private readonly List<string> _logs = new List<string>();

        private int _columns;
        private int _beams;
        private int _purlins;
        private int _walls;
        private int _roofs;

        public XuongBuilder(Document doc)
        {
            _doc = doc;
        }

        public string GetReport()
        {
            return string.Format(
                "BIMxuong2\nCot: {0}\nDam/Keo: {1}\nXa go: {2}\nTuong: {3}\nMai: {4}\n\n{5}",
                _columns, _beams, _purlins, _walls, _roofs, string.Join("\n", _logs.Take(25).ToArray()));
        }

        public void Build()
        {
            using (Transaction tx = new Transaction(_doc, "Build BIM xuong"))
            {
                tx.Start();

                SetupLevels();
                SetupGrids();
                ResolveTypes();

                CreateColumns();
                CreatePrimaryBeams();
                CreatePurlins();
                CreateWalls();
                CreateRoof();
                CreateOpenings();

                tx.Commit();
            }
        }

        private void SetupLevels()
        {
            _lvBase = FindOrCreateLevel(ModelData.BaseLevelMm, "L00_Base_+0.000");
            _lvEave = FindOrCreateLevel(ModelData.EaveLevelMm, "L01_Eave_+7.000");
            _lvRidge = FindOrCreateLevel(ModelData.RidgeLevelMm, "L02_Ridge_+11.138");
        }

        private Level FindOrCreateLevel(double elevationMm, string name)
        {
            double z = ModelData.FtFromMm(elevationMm);
            Level existing = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .FirstOrDefault(lv => Math.Abs(lv.Elevation - z) < ModelData.FtFromMm(5));

            if (existing != null)
            {
                if (existing.Name != name)
                {
                    try { existing.Name = name; } catch { }
                }
                return existing;
            }

            Level created = Level.Create(_doc, z);
            try { created.Name = name; } catch { }
            return created;
        }

        private void SetupGrids()
        {
            double yMin = ModelData.FtFromMm(-3000);
            double yMax = ModelData.FtFromMm(ModelData.BuildingLengthMm + 3000);
            double xMin = ModelData.FtFromMm(-3000);
            double xMax = ModelData.FtFromMm(ModelData.BuildingWidthMm + 3000);

            for (int i = 0; i < ModelData.GridXmm.Length; i++)
            {
                double x = ModelData.FtFromMm(ModelData.GridXmm[i]);
                EnsureGrid(ModelData.GridXNames[i], Line.CreateBound(new XYZ(x, yMin, 0), new XYZ(x, yMax, 0)));
            }

            for (int i = 0; i < ModelData.GridYmm.Length; i++)
            {
                double y = ModelData.FtFromMm(ModelData.GridYmm[i]);
                EnsureGrid(ModelData.GridYNames[i], Line.CreateBound(new XYZ(xMin, y, 0), new XYZ(xMax, y, 0)));
            }
        }

        private void EnsureGrid(string name, Line line)
        {
            Grid existing = new FilteredElementCollector(_doc)
                .OfClass(typeof(Grid))
                .Cast<Grid>()
                .FirstOrDefault(g => g.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

            if (existing != null) return;

            Grid gNew = Grid.Create(_doc, line);
            try { gNew.Name = name; } catch { }
        }

        private void ResolveTypes()
        {
            _columnSymbol = FindSymbol(BuiltInCategory.OST_StructuralColumns);
            _beamSymbol = FindSymbol(BuiltInCategory.OST_StructuralFraming);
            _purlinSymbol = FindSecondaryFramingSymbol() ?? _beamSymbol;

            _wallType = new FilteredElementCollector(_doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .FirstOrDefault();

            _roofType = new FilteredElementCollector(_doc)
                .OfClass(typeof(RoofType))
                .Cast<RoofType>()
                .FirstOrDefault();

            if (_columnSymbol != null && !_columnSymbol.IsActive) _columnSymbol.Activate();
            if (_beamSymbol != null && !_beamSymbol.IsActive) _beamSymbol.Activate();
            if (_purlinSymbol != null && !_purlinSymbol.IsActive) _purlinSymbol.Activate();
        }

        private FamilySymbol FindSymbol(BuiltInCategory bic)
        {
            return new FilteredElementCollector(_doc)
                .OfCategory(bic)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault();
        }

        private FamilySymbol FindSecondaryFramingSymbol()
        {
            List<FamilySymbol> all = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();

            FamilySymbol cType = all.FirstOrDefault(s => s.Name.IndexOf("C", StringComparison.OrdinalIgnoreCase) >= 0);
            return cType ?? all.FirstOrDefault();
        }

        private void CreateColumns()
        {
            if (_columnSymbol == null)
            {
                _logs.Add("Khong tim thay symbol cot");
                return;
            }

            for (int ix = 0; ix < ModelData.GridXmm.Length; ix++)
            {
                double x = ModelData.GridXmm[ix];
                for (int iy = 0; iy < ModelData.GridYmm.Length; iy++)
                {
                    double y = ModelData.GridYmm[iy];

                    FamilyInstance col = _doc.Create.NewFamilyInstance(
                        ModelData.PtMm(x, y, ModelData.BaseLevelMm),
                        _columnSymbol,
                        _lvBase,
                        StructuralType.Column);

                    Parameter pTop = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);
                    if (pTop != null && !pTop.IsReadOnly) pTop.Set(_lvEave.Id);

                    Parameter pOffset = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM);
                    if (pOffset != null && !pOffset.IsReadOnly)
                    {
                        pOffset.Set(ModelData.FtFromMm(ModelData.ColumnTopOffsetMm(x)));
                    }

                    SetMark(col, string.Format("KC_COL_{0}{1}", ModelData.GridXNames[ix], ModelData.GridYNames[iy]));
                    _columns++;
                }
            }
        }

        private void CreatePrimaryBeams()
        {
            if (_beamSymbol == null)
            {
                _logs.Add("Khong tim thay symbol dam");
                return;
            }

            double[] nodesX = { 0, 8000, 16000, 20000, 24000, 32000, 40000 };

            for (int iy = 0; iy < ModelData.GridYmm.Length; iy++)
            {
                double y = ModelData.GridYmm[iy];

                for (int i = 0; i < nodesX.Length - 1; i++)
                {
                    double x1 = nodesX[i];
                    double x2 = nodesX[i + 1];

                    FamilyInstance rafter = _doc.Create.NewFamilyInstance(
                        Line.CreateBound(
                            ModelData.PtMm(x1, y, ModelData.RoofZmm(x1)),
                            ModelData.PtMm(x2, y, ModelData.RoofZmm(x2))),
                        _beamSymbol,
                        _lvEave,
                        StructuralType.Beam);

                    SetStructuralUsage(rafter, StructuralInstanceUsage.Girder);
                    SetMark(rafter, string.Format("KC_RAFTER_{0}_{1}", iy + 1, i + 1));
                    _beams++;
                }
            }

            for (int ix = 0; ix < ModelData.GridXmm.Length; ix++)
            {
                double x = ModelData.GridXmm[ix];
                for (int iy = 0; iy < ModelData.GridYmm.Length - 1; iy++)
                {
                    double y1 = ModelData.GridYmm[iy];
                    double y2 = ModelData.GridYmm[iy + 1];

                    FamilyInstance tie = _doc.Create.NewFamilyInstance(
                        Line.CreateBound(
                            ModelData.PtMm(x, y1, ModelData.RoofZmm(x)),
                            ModelData.PtMm(x, y2, ModelData.RoofZmm(x))),
                        _beamSymbol,
                        _lvEave,
                        StructuralType.Beam);

                    SetStructuralUsage(tie, StructuralInstanceUsage.Joist);
                    SetMark(tie, string.Format("KC_TIE_{0}_{1}", ix + 1, iy + 1));
                    _beams++;
                }
            }
        }

        private void CreatePurlins()
        {
            if (_purlinSymbol == null)
            {
                _logs.Add("Khong tim thay symbol xa go");
                return;
            }

            IList<double> purlinXs = ModelData.BuildPurlinXmm();
            for (int i = 0; i < purlinXs.Count; i++)
            {
                double x = purlinXs[i];
                for (int iy = 0; iy < ModelData.GridYmm.Length - 1; iy++)
                {
                    double y1 = ModelData.GridYmm[iy];
                    double y2 = ModelData.GridYmm[iy + 1];

                    FamilyInstance purlin = _doc.Create.NewFamilyInstance(
                        Line.CreateBound(
                            ModelData.PtMm(x, y1, ModelData.RoofZmm(x)),
                            ModelData.PtMm(x, y2, ModelData.RoofZmm(x))),
                        _purlinSymbol,
                        _lvEave,
                        StructuralType.Beam);

                    SetStructuralUsage(purlin, StructuralInstanceUsage.Purlin);
                    SetMark(purlin, string.Format("KCP_PURLIN_X{0}_B{1}", (int)x, iy + 1));
                    _purlins++;
                }
            }
        }

        private void CreateWalls()
        {
            if (_wallType == null)
            {
                _logs.Add("Khong tim thay WallType");
                return;
            }

            double h = ModelData.FtFromMm(ModelData.WallHeightMm);

            for (int iy = 0; iy < ModelData.GridYmm.Length - 1; iy++)
            {
                MakeWall(0, ModelData.GridYmm[iy], 0, ModelData.GridYmm[iy + 1], h);
                MakeWall(ModelData.BuildingWidthMm, ModelData.GridYmm[iy], ModelData.BuildingWidthMm, ModelData.GridYmm[iy + 1], h);
            }

            for (int ix = 0; ix < ModelData.GridXmm.Length - 1; ix++)
            {
                MakeWall(ModelData.GridXmm[ix], 0, ModelData.GridXmm[ix + 1], 0, h);
                MakeWall(ModelData.GridXmm[ix], ModelData.BuildingLengthMm, ModelData.GridXmm[ix + 1], ModelData.BuildingLengthMm, h);
            }
        }

        private void MakeWall(double x1, double y1, double x2, double y2, double heightFt)
        {
            Wall wall = Wall.Create(
                _doc,
                Line.CreateBound(ModelData.PtMm(x1, y1, 0), ModelData.PtMm(x2, y2, 0)),
                _wallType.Id,
                _lvBase.Id,
                heightFt,
                0,
                false,
                false);

            SetMark(wall, string.Format("TG_{0}_{1}_{2}_{3}", (int)x1, (int)y1, (int)x2, (int)y2));
            _walls++;
        }

        private void CreateRoof()
        {
            if (_roofType == null)
            {
                _logs.Add("Khong tim thay RoofType");
                return;
            }

            CurveArray footprint = new CurveArray();
            XYZ p1 = ModelData.PtMm(0, 0, ModelData.EaveLevelMm);
            XYZ p2 = ModelData.PtMm(ModelData.BuildingWidthMm, 0, ModelData.EaveLevelMm);
            XYZ p3 = ModelData.PtMm(ModelData.BuildingWidthMm, ModelData.BuildingLengthMm, ModelData.EaveLevelMm);
            XYZ p4 = ModelData.PtMm(0, ModelData.BuildingLengthMm, ModelData.EaveLevelMm);

            footprint.Append(Line.CreateBound(p1, p2));
            footprint.Append(Line.CreateBound(p2, p3));
            footprint.Append(Line.CreateBound(p3, p4));
            footprint.Append(Line.CreateBound(p4, p1));

            ModelCurveArray mapping;
            FootPrintRoof roof = _doc.Create.NewFootPrintRoof(footprint, _lvEave, _roofType, out mapping);

            int idx = 0;
            foreach (ModelCurve mc in mapping)
            {
                if (idx == 0 || idx == 2)
                {
                    roof.set_DefinesSlope(mc, true);
                    roof.set_SlopeAngle(mc, Math.Atan(ModelData.RoofSlope));
                }
                else
                {
                    roof.set_DefinesSlope(mc, false);
                }
                idx++;
            }

            SetMark(roof, "MAI_CHINH");
            _roofs++;
        }

        private void CreateOpenings()
        {
            IList<Wall> walls = new FilteredElementCollector(_doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(w => w.LevelId == _lvBase.Id)
                .ToList();

            for (int i = 0; i < walls.Count; i++)
            {
                Wall w = walls[i];
                LocationCurve lc = w.Location as LocationCurve;
                if (lc == null || lc.Curve == null) continue;

                XYZ a = lc.Curve.GetEndPoint(0);
                XYZ b = lc.Curve.GetEndPoint(1);
                XYZ mid = (a + b) * 0.5;

                XYZ p1 = new XYZ(mid.X - ModelData.FtFromMm(600), mid.Y, ModelData.FtFromMm(900));
                XYZ p2 = new XYZ(mid.X + ModelData.FtFromMm(600), mid.Y, ModelData.FtFromMm(2100));

                try
                {
                    Opening op = _doc.Create.NewOpening(w, p1, p2);
                    SetMark(op, "CUA_MO_1200x1200");
                }
                catch
                {
                }
            }
        }

        private void SetMark(Element element, string mark)
        {
            if (element == null) return;
            Parameter p = element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
            if (p != null && !p.IsReadOnly)
            {
                p.Set(mark);
            }
        }

        private void SetStructuralUsage(FamilyInstance fi, StructuralInstanceUsage usage)
        {
            Parameter pUsage = fi.get_Parameter(BuiltInParameter.INSTANCE_STRUCT_USAGE_PARAM);
            if (pUsage != null && !pUsage.IsReadOnly)
            {
                pUsage.Set((int)usage);
            }
        }
    }
}
