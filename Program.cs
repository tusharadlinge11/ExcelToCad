
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcLine = Autodesk.AutoCAD.DatabaseServices.Line;
using AcColor = Autodesk.AutoCAD.Colors.Color;
using AcColorMethod = Autodesk.AutoCAD.Colors.ColorMethod;

[assembly: CommandClass(typeof(GraphCommand))]

public class GraphCommand
{
    static GraphCommand()
    {
        AppDomain.CurrentDomain.AssemblyResolve += (s, a) =>
        {
            string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string dll = Path.Combine(folder, new AssemblyName(a.Name).Name + ".dll");
            return File.Exists(dll) ? Assembly.LoadFrom(dll) : null;
        };
    }

    class Circuit
    {
        public string No = "";
        public string Room = "";
        public short Color;
    }

    class LoadItem
    {
        public string Code = "";
        public int Watt;
        public int Count;
    }

    [CommandMethod("GRAPH")]
    public void CreateGraph()
    {
        Document doc = AcApp.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        PromptOpenFileOptions opt = new PromptOpenFileOptions("\nSelect Excel file: ");
        opt.Filter = "Excel Files (*.xlsx)|*.xlsx";

        PromptFileNameResult res = ed.GetFileNameForOpen(opt);
        if (res.Status != PromptStatus.OK) return;

        if (!File.Exists(res.StringResult))
        {
            ed.WriteMessage("\nExcel file not found.");
            return;
        }

        Dictionary<string, List<LoadItem>> loadTable = ReadLoadData(res.StringResult);

        List<Circuit> data = new List<Circuit>
        {
           new Circuit{No="R1", Room="Living / Kitchen Lighting", Color=1},
new Circuit{No="R2", Room="Kitchen Socket", Color=1},
new Circuit{No="R3", Room="Living AC", Color=1},
new Circuit{No="R4", Room="Refrigerator", Color=1},

new Circuit{No="Y1", Room="C.Bedroom / C.Toilet Lighting", Color=2},
new Circuit{No="Y2", Room="C.Bedroom AC", Color=2},
new Circuit{No="Y3", Room="C.Toilet Geyser", Color=2},
new Circuit{No="Y4", Room="Micro Wave / Mixer", Color=2},

new Circuit{No="B1", Room="M.Bedroom / M.Toilet Lighting", Color=5},
new Circuit{No="B2", Room="M.Bedroom AC", Color=5},
new Circuit{No="B3", Room="M.Toilet Geyser", Color=5},
new Circuit{No="B4", Room="Spare", Color=5}
        };

        using (doc.LockDocument())
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord ms = (BlockTableRecord)tr.GetObject(
                bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            CreateLayer(db, tr, "SLD_RED", 1);
            CreateLayer(db, tr, "SLD_YELLOW", 2);
            CreateLayer(db, tr, "SLD_BLUE", 5);
            CreateLayer(db, tr, "SLD_TEXT", 7);
            CreateLayer(db, tr, "SLD_FRAME", 8);
            CreateLayer(db, tr, "SLD_GREEN", 3);
            CreateLayer(db, tr, "SLD_MAGENTA", 6);

            Text(ms, tr, "3BHK FLAT TYPE 1", 210, 530, 18, "SLD_TEXT");

            DrawMainPanel(ms, tr);
            DrawAllCircuits(ms, tr, data, loadTable);

            tr.Commit();
        }

        ed.WriteMessage("\nGraph created successfully.");
    }

    void DrawMainPanel(BlockTableRecord ms, Transaction tr)
    {
        Rect(ms, tr, 170, 430, 560, 500, "SLD_BLUE");

        Text(ms, tr, "SAMADHAN GORAI", 175, 490, 3.2, "SLD_TEXT");
        Text(ms, tr, "3BHK FLAT TYPE-1", 175, 482, 3.2, "SLD_TEXT");
        Text(ms, tr, "6 WAY TPN DB", 175, 474, 3.2, "SLD_TEXT");

        Text(ms, tr, "MAIN DB", 340, 488, 4.2, "SLD_TEXT");
        Text(ms, tr, "25A 4P MCB", 335, 475, 3.2, "SLD_TEXT");
        Text(ms, tr, "25A 4P 30mA RCCB", 328, 465, 3.2, "SLD_TEXT");

        Text(ms, tr, "R PHASE : 1.5KW", 510, 490, 3.0, "SLD_TEXT");
        Text(ms, tr, "Y PHASE : 1.5KW", 510, 482, 3.0, "SLD_TEXT");
        Text(ms, tr, "B PHASE : 1.5KW", 510, 474, 3.0, "SLD_TEXT");

        double[] xs = { 230, 242, 254, 266, 290, 302, 314, 326, 370, 382, 394, 406 };
        string[] nos = { "R1", "R2", "R3", "R4", "Y1", "Y2", "Y3", "Y4", "B1", "B2", "B3", "B4" };

        for (int i = 0; i < xs.Length; i++)
        {
            string layer = i < 4 ? "SLD_RED" : i < 8 ? "SLD_YELLOW" : "SLD_BLUE";
            Text(ms, tr, nos[i], xs[i] - 3, 445, 3.2, layer);
            Line(ms, tr, xs[i], 430, xs[i], 420, layer);
        }
    }

    void DrawAllCircuits(BlockTableRecord ms, Transaction tr, List<Circuit> data,
     Dictionary<string, List<LoadItem>> loadTable)
    {
        double topY = 430;

        Dictionary<string, double> xMap = new Dictionary<string, double>
    {
        {"R1",230}, {"R2",242}, {"R3",254}, {"R4",266},
        {"Y1",290}, {"Y2",302}, {"Y3",314}, {"Y4",326},
        {"B1",370}, {"B2",382}, {"B3",394}, {"B4",406}
    };

        double leftY = 390;
        double midY = 0;
        double rightY = 390;

        double smallGap = 10;

        foreach (Circuit c in data)
        {
            double x = xMap[c.No];
            string layer = GetLayer(c.Color);

            bool isR = c.No.StartsWith("R");
            bool isY = c.No.StartsWith("Y");
            bool isB = c.No.StartsWith("B");

            if (isY && midY == 0)
                midY = leftY - 40;   // R series पूर्ण झाल्यावर Y series खाली सुरू

            double y = isR ? leftY : isY ? midY : rightY;

            if (c.No == "B1") y = 250;
            else if (c.No == "B2") y = 285;
            else if (c.No == "B3") y = 320;
            else if (c.No == "B4") y = 355;

            List<LoadItem> circuitLoads = null;

            if (loadTable.TryGetValue(c.No, out circuitLoads) == false)
                circuitLoads = null;

            int branchTotalForLine = 1;

            if (circuitLoads != null && circuitLoads.Count > 0)
                branchTotalForLine = circuitLoads.Sum(a => Math.Min(15, Math.Max(1, a.Count)));

            double lastEntityY =
                y - (branchTotalForLine * 3.5) -
                ((circuitLoads == null ? 1 : circuitLoads.Count) * 10.0);

            Line(ms, tr, x, topY, x, lastEntityY, layer);
            Circle(ms, tr, x, y, 1.6, layer);

            if (isR || isY)
            {
                double listCenterY =
                    y - ((branchTotalForLine * 3.5) / 2.0);

                if (circuitLoads != null && circuitLoads.Count > 0)
                {
                    DrawEntities(ms, tr,
                        circuitLoads,
                        x - 10,
                        y,
                        x,
                        layer,
                        true);
                }

                Text(ms, tr,
                    c.Room,
                    x - 120,
                    listCenterY - 2,
                    3.5,
                    "SLD_TEXT");
            }

            if (isB)
            {
                double listCenterY =
                    y - ((branchTotalForLine * 3.5) / 2.0);

                if (circuitLoads != null && circuitLoads.Count > 0)
                {
                    DrawEntities(ms, tr,
                        circuitLoads,
                        x + 25,
                        y,
                        x,
                        layer,
                        false);
                }

                Text(ms, tr,
                    c.Room,
                    x + 75,
                    listCenterY - 2,
                    3.5,
                    "SLD_TEXT");
            }

            double used =
                branchTotalForLine * 3.5 +
                ((circuitLoads == null ? 1 : circuitLoads.Count) * 10.0) +
                smallGap;

            if (c.No == "R4") used += 40;
            if (c.No == "Y4") used += 40;

            if (isR) leftY -= used;
            if (isY) midY -= used;
            if (isB) rightY -= used;
        }
    }
    void DrawEntities(BlockTableRecord ms, Transaction tr, List<LoadItem> loads,
    double textX, double baseY, double busX, string layer, bool leftSide)
    {
        if (loads == null || loads.Count == 0)
            return;

        double y = baseY - 8;

        double lineLength = 18;
        double branchLength = 10;
        double branchGap = 3.5;
        double entityGap = 10.0;

        foreach (LoadItem item in loads)
        {
            int branchCount = item.Count;

            if (branchCount < 1)
                branchCount = 1;

            if (branchCount > 15)
                branchCount = 15;

            double endX = leftSide
                ? busX - lineLength
                : busX + lineLength;

            // main entity connect line
            Line(ms, tr,
                busX,
                y + 1.2,
                endX,
                y + 1.2,
                layer);

            // main connection dot
            Circle(ms, tr,
                busX,
                y + 1.2,
                1.1,
                layer);

            double branchX;
            double labelX;

            if (leftSide)
            {
                // R/Y branches right side ने connect
                branchX = endX + branchLength;

                // R/Y branch name line पूर्ण झाल्यावर left side ला
                labelX = branchX - 45;
            }
            else
            {
                // B branches left side ने connect
                branchX = endX - branchLength;

                // B branch name line पूर्ण झाल्यावर right side ला
                labelX = endX + 5;
            }

            double firstY = y + 1.2;
            double lastY = y - ((branchCount - 1) * branchGap) + 1.2;

            // vertical connector for branches
            Line(ms, tr,
                branchX,
                firstY,
                branchX,
                lastY,
                layer);

            for (int b = 0; b < branchCount; b++)
            {
                double by = y - (b * branchGap) + 1.2;

                // horizontal sub-branch
                Line(ms, tr,
                    endX,
                    by,
                    branchX,
                    by,
                    layer);

                // full entity names
                string fullName = item.Code;

                if (fullName == "TL") fullName = "Tube Light";
                else if (fullName == "BL") fullName = "Bracket Light";
                else if (fullName == "CF") fullName = "Ceiling Fan";
                else if (fullName == "CL") fullName = "Ceiling Light";
                else if (fullName == "BB") fullName = "Bell";
                else if (fullName == "EF") fullName = "Exhaust Fan";
                else if (fullName == "FL") fullName = "Foot Light";
                else if (fullName == "CTV") fullName = "Computer/TV";
                else if (fullName == "CHG") fullName = "Charger";
                else if (fullName == "SS") fullName = "Shaver Socket";
                else if (fullName == "SO") fullName = "Socket";
                else if (fullName == "WM") fullName = "Washing Machine";
                else if (fullName == "WP") fullName = "Water Purifier";
                else if (fullName == "CP") fullName = "Chimney Point";
                else if (fullName == "MX") fullName = "Mixer";
                else if (fullName == "MW") fullName = "Microwave";
                else if (fullName == "REF") fullName = "Refrigerator";
                else if (fullName == "AC") fullName = "Air Conditioner";
                else if (fullName == "GY") fullName = "Geyser";

                string branchName = fullName + " " + (b + 1);

                // branch name
                Text(ms, tr,
                    branchName,
                    labelX,
                    by - 1.2,
                    2.5,
                    "SLD_GREEN");
            }

            // next entity position
            y -= (branchCount * branchGap) + entityGap;
        }
    }

    Dictionary<string, List<LoadItem>> ReadLoadData(string path)
    {
        Dictionary<string, List<LoadItem>> dict = new Dictionary<string, List<LoadItem>>();

        using (XLWorkbook wb = new XLWorkbook(path))
        {
            IXLWorksheet ws = wb.Worksheets.First();

            int headerRow = 8;
            int wattRow = 9;

            for (int r = 10; r <= 30; r++)
            {
                string refNo = ws.Cell(r, 4).GetString().Trim();

                if (string.IsNullOrWhiteSpace(refNo)) continue;
                if (!(refNo.StartsWith("R") || refNo.StartsWith("Y") || refNo.StartsWith("B"))) continue;

                List<LoadItem> list = new List<LoadItem>();

                for (int c = 9; c <= 28; c++)
                {
                    string header = ws.Cell(headerRow, c).GetString().Trim();

                    if (string.IsNullOrWhiteSpace(header))
                        header = ws.Cell(7, c).GetString().Trim();

                    string code = MapHeaderToCode(header);
                    int watt = ReadInt(ws.Cell(wattRow, c));
                    int count = ReadInt(ws.Cell(r, c));

                    if (count > 0 && watt > 0 && code != "")
                    {
                        list.Add(new LoadItem
                        {
                            Code = code,
                            Watt = watt,
                            Count = count
                        });
                    }
                }

                if (list.Count > 0)
                    dict[refNo] = list;
            }
        }

        return dict;
    }

    int ReadInt(IXLCell cell)
    {
        try
        {
            return (int)Math.Round(cell.GetDouble());
        }
        catch
        {
            string s = cell.GetString();
            string clean = new string(s.Where(ch => char.IsDigit(ch) || ch == '.').ToArray());

            if (double.TryParse(clean, out double d))
                return (int)Math.Round(d);

            return 0;
        }
    }

    string MapHeaderToCode(string header)
    {
        if (string.IsNullOrWhiteSpace(header)) return "";

        string h = header.ToUpperInvariant();

        if (h.Contains("TUBE")) return "TL";
        if (h.Contains("BRACKET")) return "BL";
        if (h.Contains("CEILING") && h.Contains("FAN")) return "CF";
        if (h.Contains("CEILING") && h.Contains("LIGHT")) return "CL";
        if (h.Contains("CHANDEL")) return "CL";
        if (h.Contains("BELL")) return "BB";
        if (h.Contains("EXHAUST")) return "EF";
        if (h.Contains("FOOT")) return "FL";
        if (h.Contains("COMPUTER") || h.Contains("TV")) return "CTV";
        if (h.Contains("CHARGER")) return "CHG";
        if (h.Contains("SHAVER")) return "SS";
        if (h.Contains("SOCKET")) return "SO";
        if (h.Contains("WASH")) return "WM";
        if (h.Contains("WATER")) return "WP";
        if (h.Contains("CHIMNEY")) return "CP";
        if (h.Contains("MIXER")) return "MX";
        if (h.Contains("MICRO")) return "MW";
        if (h.Contains("REFRIGE")) return "REF";
        if (h.Contains("AC")) return "AC";
        if (h.Contains("GEYSER")) return "GY";

        return header.Length > 4
            ? header.Substring(0, 4).ToUpperInvariant()
            : header.ToUpperInvariant();
    }

    string GetLayer(short color)
    {
        if (color == 1) return "SLD_RED";
        if (color == 2) return "SLD_YELLOW";
        return "SLD_BLUE";
    }

    void Line(BlockTableRecord ms, Transaction tr,
        double x1, double y1, double x2, double y2, string layer)
    {
        AcLine l = new AcLine(
            new Point3d(x1, y1, 0),
            new Point3d(x2, y2, 0));

        l.Layer = layer;
        ms.AppendEntity(l);
        tr.AddNewlyCreatedDBObject(l, true);
    }

    void Circle(BlockTableRecord ms, Transaction tr,
        double x, double y, double r, string layer)
    {
        Circle c = new Circle(new Point3d(x, y, 0), Vector3d.ZAxis, r);
        c.Layer = layer;

        ms.AppendEntity(c);
        tr.AddNewlyCreatedDBObject(c, true);
    }

    void Rect(BlockTableRecord ms, Transaction tr,
        double x1, double y1, double x2, double y2, string layer)
    {
        Line(ms, tr, x1, y1, x2, y1, layer);
        Line(ms, tr, x2, y1, x2, y2, layer);
        Line(ms, tr, x2, y2, x1, y2, layer);
        Line(ms, tr, x1, y2, x1, y1, layer);
    }

    void Text(BlockTableRecord ms, Transaction tr,
        string value, double x, double y, double h, string layer)
    {
        DBText t = new DBText();
        t.Position = new Point3d(x, y, 0);
        t.TextString = value ?? "";
        t.Height = h;
        t.Layer = layer;

        ms.AppendEntity(t);
        tr.AddNewlyCreatedDBObject(t, true);
    }

    void CreateLayer(Database db, Transaction tr, string name, short color)
    {
        LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

        if (!lt.Has(name))
        {
            lt.UpgradeOpen();

            LayerTableRecord ltr = new LayerTableRecord();
            ltr.Name = name;
            ltr.Color = AcColor.FromColorIndex(AcColorMethod.ByAci, color);

            lt.Add(ltr);
            tr.AddNewlyCreatedDBObject(ltr, true);
        }
    }
}