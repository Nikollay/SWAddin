using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Windows.Forms;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SWAddin
{
    class Board
    {
        public List<Component> components;
        public List<Circle> circles;
        //public List<Point> point;
        public List<Object> sketh, cutout;
        public double thickness;
        public int ver;

        public static Board GetfromXML(string filename)
        {
            Board board = new Board();
            XDocument doc=GetXML(filename, board);
            if (doc == null) { return null; }
            IEnumerable<XElement> elements = doc.Root.Elements();
            board.components = new List<Component>();
            foreach (XElement e in elements)
            {
                if (e.FirstAttribute.Name == "AD_ID") { GetBodyfromXElement(e, board); }
                if (e.FirstAttribute.Name == "ID") { board.components.Add(GetfromXElement(e)); }
            }
            return board;
        }
        private static XDocument GetXML(string filename, Board board)
        {
            XDocument doc = XDocument.Load(filename);
            XElement element = (XElement)doc.Root.FirstNode;
            switch (element.Name.ToString())
            {
                case "transactions":
                    board.ver = 1;
                    return GetXML1(filename);   
                case "transaction":
                     if (element.Attribute("Type").Value == "SOLIDWORKS") 
                    {
                        board.ver = 3; 
                        return null;// GetXML3(filename); 
                    }
                    else
                    {
                        board.ver = 2;
                        return GetXML2(filename);
                    }
                default:
                    return null;
            }
         }
        private static XDocument GetXML1(string filename)
        {
            XAttribute a1, a2;
            XDocument doc_out, doc = XDocument.Load(filename);
            XElement componentXML, atribute, XML;
            IEnumerable<XElement> elements2, elements1;
            doc_out = new XDocument();
            XML = new XElement("XML");
            doc_out.Add(XML);
            try
            {
                elements1 = doc.Root.Element("transactions").Element("transaction").Element("document").Element("configuration").Element("references").Elements();
              
                foreach (XElement e1 in elements1)
                {
                    elements2 = e1.Element("configuration").Elements();
                    if (e1.FirstAttribute.Value == "Документация") { continue; }
                    componentXML = new XElement("componentXML", new XAttribute(e1.FirstAttribute.Name, e1.FirstAttribute.Value));

                    foreach (XElement e2 in elements2)
                    {
                        a1 = e2.Attribute("name");
                        a2 = e2.Attribute("value");
                        if (a1.Value == "Раздел_Сп" | a1.Value == "Fitted" | a1.Value == "GUID") { continue; }
                        atribute = new XElement("attribute", a1, a2);
                        componentXML.Add(atribute);
                    }
                    XML.Add(componentXML);
                }
            }
            catch
            {
                return null;
            }
          
         return doc_out;
            
        }
        private static XDocument GetXML2(string filename)
        {
            string descriptionPCB = "";
            string comnpName = "";
            XAttribute a1, a2;
            XDocument doc_out, doc = XDocument.Load(filename);
            XElement componentXML, atribute, XML, tmpXEl;
            IEnumerable<XElement> elements, elements2;
            doc_out = new XDocument();
            XML = new XElement("XML");
            doc_out.Add(XML);
            try
            {
                elements = doc.Root.Element("transaction").Element("project").Element("configurations").Element("configuration").Element("graphs").Elements();
                tmpXEl = elements.First(item => item.Attribute("name").Value.Equals("Обозначение_PCB"));
                descriptionPCB = tmpXEl.Attribute("value").Value;
                componentXML = new XElement("componentXML", new XAttribute("AD_ID", descriptionPCB));
                elements = doc.Root.Element("transaction").Element("project").Element("configurations").Element("configuration").Element("componentsPCB").Element("component_pcb").Element("properties").Elements();
                foreach (XElement e in elements)
                {
                    a1 = e.Attribute("name");
                    a2 = e.Attribute("value");
                    //if (a1.Value == "Раздел_Сп" | a1.Value == "Fitted" | a1.Value == "GUID") { continue; }
                    atribute = new XElement("attribute", a1, a2);
                    componentXML.Add(atribute);
                }
                XML.Add(componentXML);
                elements = doc.Root.Element("transaction").Element("project").Element("configurations").Element("configuration").Element("components").Elements();
                foreach (XElement e in elements)
                {
                    elements2 = e.Element("properties").Elements();
                    tmpXEl = elements2.First(item => item.Attribute("name").Value.Equals("Наименование"));
                    comnpName = tmpXEl.Attribute("value").Value;
                    componentXML = new XElement("componentXML", new XAttribute("ID", comnpName));
                    foreach (XElement e2 in elements2)
                    {
                        atribute = new XElement("attribute", e2.Attribute("name"), e2.Attribute("value"));
                        componentXML.Add(atribute);
                    }
                    XML.Add(componentXML);
                }
            }
            catch
            {
                return null;
            }
            return doc_out;
        }
        private static XDocument GetXML3(string filename)
        {
            string comnpName = "";
            XDocument doc_out, doc = XDocument.Load(filename);
            XElement componentXML, atribute, XML, tmpXEl;
            IEnumerable<XElement> elements, elements2;
            doc_out = new XDocument();
            XML = new XElement("XML");
            doc_out.Add(XML);
            try
            {
                elements = doc.Root.Element("transaction").Element("project").Element("configurations").Element("configuration").Element("components").Elements();
                foreach (XElement e in elements)
                {
                    elements2 = e.Element("properties").Elements();
                    tmpXEl = elements2.First(item => item.Attribute("name").Value.Equals("Наименование"));
                    comnpName = tmpXEl.Attribute("value").Value;
                    componentXML = new XElement("componentXML", new XAttribute("ID", comnpName));
                    foreach (XElement e2 in elements2)
                    {
                        atribute = new XElement("attribute", e2.Attribute("name"), e2.Attribute("value"));
                        componentXML.Add(atribute);
                    }
                    XML.Add(componentXML);
                }
            }
            catch
            {
                return null;
            }
            //doc_out.Save("d:\\Домашняя работа\\test.xml");
            return doc_out;
        }
        private static Component GetfromXElement(XElement el)
        {
            IEnumerable<XElement> elements = el.Elements();
            Component component = new Component();
            foreach (XElement e in elements)
            {
                switch (e.Attribute("name").Value)
                {
                    case "DM_PhysicalDesignator":
                        component.physicalDesignator = e.Attribute("value").Value;
                        break;
                    case "Позиционное обозначение":
                        component.physicalDesignator = e.Attribute("value").Value;
                        break;
                    case "Наименование":
                        component.title = e.Attribute("value").Value;
                        break;
                    case "Footprint":
                        component.footprint = e.Attribute("value").Value;
                        break;
                    case "Part Number":
                        component.part_Number = e.Attribute("value").Value;
                        break;
                    case "X":
                        component.x = double.Parse(e.Attribute("value").Value) / 1000;
                        break;
                    case "Y":
                        component.y = double.Parse(e.Attribute("value").Value) / 1000;
                        break;
                    case "Z":
                        component.z = double.Parse(e.Attribute("value").Value) / 1000;
                        break;
                    case "Layer":
                        component.layer = int.Parse(e.Attribute("value").Value);
                        break;
                    case "Rotation":
                        component.rotation = double.Parse(e.Attribute("value").Value);
                        break;
                    case "StandOff":
                        component.standOff = double.Parse(e.Attribute("value").Value) / 1000;
                        break;
                }
            }
            return component;
        }
        private static void GetBodyfromXElement(XElement el, Board board)
        {
            IEnumerable<XElement> elements = el.Elements();
            string str;
            string[] strsplit, strToDbl;
            Line skLine;
            Arc skArc;
            Circle skCircle;
            board.sketh = new List<object>();
            board.cutout = new List<object>();
            board.circles = new List<Circle>();
            foreach (XElement e in elements)
            {
                switch (e.Attribute("name").Value)
                {
                    case "BOARD_OUTLINE":
                        str = e.Attribute("value").Value;
                        strsplit = str.Split((char)35);
                        for (int i = 0; i < strsplit.Length; i++)
                        {
                            strToDbl = strsplit[i].Split((char)59);
                            if (strToDbl.Length == 4)
                            {
                                skLine = new Line();
                                skLine.x1 = double.Parse(strToDbl[0]) / 1000;
                                skLine.y1 = double.Parse(strToDbl[1]) / 1000;
                                skLine.x2 = double.Parse(strToDbl[2]) / 1000;
                                skLine.y2 = double.Parse(strToDbl[3]) / 1000;
                                board.sketh.Add(skLine);
                            }
                            if (strToDbl.Length == 9)
                            {
                                skArc = new Arc();
                                skArc.x1 = double.Parse(strToDbl[0]) / 1000;
                                skArc.y1 = double.Parse(strToDbl[1]) / 1000;
                                skArc.x2 = double.Parse(strToDbl[2]) / 1000;
                                skArc.y2 = double.Parse(strToDbl[3]) / 1000;
                                skArc.xc = double.Parse(strToDbl[7]) / 1000;
                                skArc.yc = double.Parse(strToDbl[8]) / 1000;
                                skArc.direction = 0;
                                board.sketh.Add(skArc);
                            }
                        }
                        break;
                    case "BOARD_CUTOUT":
                        str = e.Attribute("value").Value;
                        strsplit = str.Split((char)35);
                        for (int i = 0; i < strsplit.Length; i++)
                        {
                            strToDbl = strsplit[i].Split((char)59);
                            if (strToDbl.Length == 4)
                            {
                                skLine = new Line();
                                skLine.x1 = double.Parse(strToDbl[0]) / 1000;
                                skLine.y1 = double.Parse(strToDbl[1]) / 1000;
                                skLine.x2 = double.Parse(strToDbl[2]) / 1000;
                                skLine.y2 = double.Parse(strToDbl[3]) / 1000;
                                board.cutout.Add(skLine);
                            }
                            if (strToDbl.Length == 9)
                            {
                                skArc = new Arc();
                                skArc.x1 = double.Parse(strToDbl[0]) / 1000;
                                skArc.y1 = double.Parse(strToDbl[1]) / 1000;
                                skArc.x2 = double.Parse(strToDbl[2]) / 1000;
                                skArc.y2 = double.Parse(strToDbl[3]) / 1000;
                                skArc.xc = double.Parse(strToDbl[7]) / 1000;
                                skArc.yc = double.Parse(strToDbl[8]) / 1000;
                                skArc.direction = 1;
                                board.cutout.Add(skArc);
                            }
                        }
                        break;
                    case "DRILLED_HOLES":
                        str = e.Attribute("value").Value;
                        strsplit = str.Split((char)35);
                        for (int i = 0; i < strsplit.Length; i++)
                        {
                            strToDbl = strsplit[i].Split((char)59);
                            if (strToDbl.Length == 4)
                            {
                                skCircle = new Circle();
                                skCircle.radius = double.Parse(strToDbl[0]) / 2000;
                                skCircle.xc = double.Parse(strToDbl[1]) / 1000;
                                skCircle.yc = double.Parse(strToDbl[2]) / 1000;
                                board.circles.Add(skCircle);
                            }
                        }
                        break;
                    case "Толщина, мм":
                        board.thickness = double.Parse(e.Attribute("value").Value) / 1000;
                        break;
                }
            }
        }    
        public static string GetFilename()
        {
            string filename;
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.InitialDirectory = "D:\\Домашняя работа";
            fileDialog.Filter = "brd files (*.brd)|*.brd|xml files (*.xml)|*.xml";
            fileDialog.FilterIndex = 2;
            fileDialog.RestoreDirectory = true;
            fileDialog.Title = "Открыть файл";
            fileDialog.ShowDialog();
            filename = fileDialog.FileName;
            if (String.IsNullOrWhiteSpace(filename)) { System.Environment.Exit(0); }
            return filename;
        }

        public static Excel.Workbook GetfromXDocument(XDocument doc, Excel.Application xlApp)
        {
            IEnumerable<XElement> elements1, elements2;
            Excel.Worksheet wh;
            Excel.Workbook wb = xlApp.Workbooks.Add("D:\\PDM\\EPDM_LIBRARY\\EPDM_Specification\\sp.xls");
            XElement tmpXEl;
            string designation;
            //Заполняем шапку
            wh = wb.Worksheets[1];
            elements1 = doc.Root.Element("transaction").Element("project").Element("configurations").Element("configuration").Element("graphs").Elements();
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("Проект"));
            wh.Cells[1, 1] = tmpXEl.Attribute("value").Value;
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("Перв.Примен."));
            wh.Cells[1, 3] = tmpXEl.Attribute("value").Value;
            wh.Cells[3, 14] = "Документация";
            wh.Cells[3, 14].Font.Underline = true;
            wh.Cells[3, 14].Font.Bold = true;
            wh.Cells[3, 14].HorizontalAlignment = -4108; // xlCenter
            wh.Cells[3, 14].VerticalAlignment = -4108; // xlCenter
            wh.Cells[5, 4] = "A3";
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("Обозначение"));
            designation = tmpXEl.Attribute("value").Value;
            wh.Cells[5, 9] = tmpXEl.Attribute("value").Value + "СБ";
            wh.Cells[5, 14] = "Сборочный чертеж";
            wh.Cells[32, 12] = tmpXEl.Attribute("value").Value;
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("Наименование"));
            wh.Cells[35, 12] = tmpXEl.Attribute("value").Value;
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("п_Разраб"));
            wh.Cells[35, 8] = tmpXEl.Attribute("value").Value;
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("п_Пров_P"));
            wh.Cells[36, 8] = tmpXEl.Attribute("value").Value;
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("п_Н_контр"));
            wh.Cells[38, 8] = tmpXEl.Attribute("value").Value;
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("п_Утв"));
            wh.Cells[39, 8] = tmpXEl.Attribute("value").Value;

            //Заполняем словарь
            elements1 = doc.Root.Element("transaction").Element("project").Element("configurations").Element("configuration").Element("components").Elements();
            Record component;
            Dictionary<string, Record> dict;

            SortedDictionary<string, Record> dDocumentation, dAssembly, dParts, dStandard, dOther, dMaterials, dKits, dNone;
            dDocumentation = new SortedDictionary<string, Record>();
            dAssembly = new SortedDictionary<string, Record>();
            dParts = new SortedDictionary<string, Record>();
            dStandard = new SortedDictionary<string, Record>();
            dOther = new SortedDictionary<string, Record>();
            dMaterials = new SortedDictionary<string, Record>();
            dKits = new SortedDictionary<string, Record>();
            dNone = new SortedDictionary<string, Record>();

            dict = new Dictionary<string, Record>();
            string key;
            foreach (XElement e1 in elements1)
            {
                component = new Record();
                component.quantity = 1;
                elements2 = e1.Element("properties").Elements();
                foreach (XElement e2 in elements2)
                {
                    switch (e2.Attribute("name").Value)
                    {
                        case "Формат":
                            component.format = e2.Attribute("value").Value;
                            break;
                        case "Обозначение":
                            component.designation = e2.Attribute("value").Value;
                            break;
                        case "Наименование":
                            component.title = e2.Attribute("value").Value;
                            break;
                        case "Примечание":
                            component.note = e2.Attribute("value").Value;
                            break;
                        case "Раздел СП":
                            component.chapter = e2.Attribute("value").Value;
                            break;
                    }
                }
                key = component.designation + (char)32 + component.title;
                if (!dict.ContainsKey(key)) { dict.Add(key, component); }
                else dict[key].quantity++;
            }
            //Заполнили словарь *******
            //Сортировка
            dict.OrderBy(k => k.Key).GroupBy(g => g.Value.chapter);

            string partition = "Документация";
            int j = 6;

            //Заполняем листы
            foreach (KeyValuePair<string, Record> d in dict)
            {
                if ((j % 4) == 0) { j++; }
                if (!d.Value.chapter.Equals(partition))
                {
                    wh.Cells[j + 2, 14] = d.Value.chapter;
                    wh.Cells[j + 2, 14].Font.Underline = true;
                    wh.Cells[j + 2, 14].Font.Bold = true;
                    wh.Cells[j + 2, 14].HorizontalAlignment = -4108; //xlCenter
                    wh.Cells[j + 2, 14].VerticalAlignment = -4108; //xlCenter
                    j += 5;
                    partition = d.Value.chapter;
                }

                if (j > 26 & wh.Name.Equals(1))
                {
                    wb.Sheets.get_Item(wb.Worksheets.Count - 1).Copy(wb.Sheets.get_Item(wb.Worksheets.Count - 2));
                    wh = wb.Sheets.get_Item(wb.Worksheets.Count - 2);
                    j = 4;
                }

                if (j > 33)
                {
                    wb.Sheets.get_Item(wb.Worksheets.Count - 1).Copy(wb.Sheets.get_Item(wb.Worksheets.Count - 2));
                    wh = wb.Sheets.get_Item(wb.Worksheets.Count - 2);
                    j = 4;
                }

                wh.Cells[j, 4] = d.Value.format;
                wh.Cells[j, 9] = d.Value.designation;
                wh.Cells[j, 20] = d.Value.quantity;
                wh.Cells[j, 21] = d.Value.note;

                if (d.Value.title.Length < 32) { wh.Cells[j, 14] = d.Value.title; }

                if (d.Value.title.Length > 31)
                {
                    wh.Cells[j, 14] = d.Value.title.Substring(0, 31);
                    wh.Cells[j + 1, 14] = d.Value.title.Substring(31);
                    j += 1;
                }

            }
            //Заполнили
            wb.Sheets.get_Item(wb.Worksheets.Count - 1).Delete();//Удаляем лист шаблон

            if (wb.Worksheets.Count == 2)
            {
                wh = wb.Sheets.get_Item(1);
                wh.Cells[36, 19] = "";
            }
            if (wb.Worksheets.Count < 4) { wb.Sheets.get_Item("ЛРИ").Delete(); } //Удаляем лист ЛРИ
                wh = wb.Sheets.get_Item(1);
                wh.Cells[36, 22] = wb.Worksheets.Count;

                for (int i = 2; i < wb.Worksheets.Count; i++)
                {
                    wh = wb.Sheets.get_Item(i);
                    wh.Cells[35, 12] = designation;
                    if (!wh.Name.Equals("ЛРИ"))
                    {
                        wh.Name = i.ToString();
                        wh.Cells[37, 22] = i;
                    }
                    if (wh.Name.Equals("ЛРИ")) { wh.Cells[37, 19] = wb.Worksheets.Count; }
                }
            
        return wb;
        }
        private static XDocument GetBRD(string filename)
        {

            //string[] line = File.ReadAllLines(filename, System.Text.Encoding.GetEncoding(1251));
            //List<string> list = new List<string>(line);
            //int board_outline_start = 0, drilled_holes_start = 0, placement_start = 0, board_outline_end = 0, drilled_holes_end = 0, placement_end = 0;

            //List<string> board_outline;
            //List<string> drilled_holes;
            //List<string> placement;
            //string[] strSplit;

            //for (int i = 0; i < list.Count; i++)
            //{
            //    if (list[i].ToUpper().Contains(".BOARD_OUTLINE UNOWNED")) { board_outline_start = i; }
            //    if (list[i].ToUpper().Contains(".END_BOARD_OUTLINE")) { board_outline_end = i; }
            //    if (list[i].ToUpper().Contains(".DRILLED_HOLES")) { drilled_holes_start = i; }
            //    if (list[i].ToUpper().Contains(".END_DRILLED_HOLES")) { drilled_holes_end = i; }
            //    if (list[i].ToUpper().Contains(".PLACEMENT")) { placement_start = i; }
            //    if (list[i].ToUpper().Contains(".END_PLACEMENT")) { placement_end = i; }
            //}

            //board_outline = list.GetRange(board_outline_start + 1, board_outline_end - board_outline_start - 1);
            //drilled_holes = list.GetRange(drilled_holes_start + 1, drilled_holes_end - drilled_holes_start - 1);
            //placement = list.GetRange(placement_start + 1, placement_end - placement_start - 1);
            ////for (int i = 0; i < placement.Count; i++) { Console.WriteLine(placement[i]); }
            //board.thickness = double.Parse(board_outline[0].Replace(".", ",")) / 1000;
            //board_outline.RemoveAt(0);


            //board.point = new List<Point>();
            //for (int i = 0; i < board_outline.Count; i++)
            //{
            //    point = new Point();
            //    strSplit = board_outline[i].Split((char)32);
            //    point.x = float.Parse(strSplit[1].Replace(".", ",")) / 1000;
            //    point.y = float.Parse(strSplit[2].Replace(".", ",")) / 1000;
            //    point.angle = float.Parse(strSplit[3].Replace(".", ","));
            //    board.point.Add(point);
            //}

            //board.circles = new List<Circle>();
            //for (int i = 0; i < drilled_holes.Count; i++)
            //{
            //    circle = new Circle();
            //    strSplit = drilled_holes[i].Split((char)32);
            //    circle.xc = float.Parse(strSplit[1].Replace(".", ",")) / 1000;
            //    circle.yc = float.Parse(strSplit[2].Replace(".", ",")) / 1000;
            //    circle.radius = float.Parse(strSplit[0].Replace(".", ",")) / 2000;
            //    if (!strSplit[5].Contains("VIA")) { board.circles.Add(circle); }
            //}

            //board.components = new List<Component>();
            //for (int i = 0; i < placement.Count; i++)
            //{
            //    if (i % 2 == 0)
            //    {
            //        component = new Component();
            //        strSplit = placement[i].Split((char)32);
            //        component.footprint = strSplit[0];
            //        component.physicalDesignator = strSplit[strSplit.Length - 1];
            //        component.part_Number = placement[i].Replace(component.footprint, "");
            //        component.part_Number = component.part_Number.Replace(component.physicalDesignator, "");
            //        component.part_Number = component.part_Number.Trim().Trim('\"');

            //        strSplit = placement[i + 1].Split((char)32);
            //        component.x = float.Parse(strSplit[0].Replace(".", ",")) / 1000;
            //        component.y = float.Parse(strSplit[1].Replace(".", ",")) / 1000;
            //        component.z = board.thickness;
            //        component.standOff = float.Parse(strSplit[2].Replace(".", ",")) / 1000;
            //        component.rotation = float.Parse(strSplit[3].Replace(".", ","));
            //        switch (strSplit[4])
            //        {
            //            case "TOP":
            //                component.layer = 1;
            //                break;
            //            case "BOTTOM":
            //                component.layer = 0;
            //                break;
            //        }
            //        board.components.Add(component);
            //        //Console.WriteLine(component.part_Number);
            //        //Console.WriteLine(component.x);
            //        //Console.WriteLine(component.y);
            //        //Console.WriteLine(component.rotation);
            //    }
            //}
            //Console.ReadKey();

            ////Эскизы
            //for (int i = 1; i < board.point.Count; i++)
            //{
            //    swModel.SketchManager.CreateLine(board.point[i - 1].x, board.point[i - 1].y, 0, board.point[i].x, board.point[i].y, 0);
            //}
            //swModel.FeatureManager.FeatureExtrusion3(true, false, false, 0, 0, board.thickness, board.thickness, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);
            //swModel.ClearSelection2(true);

            //plane.Select2(false, -1);
            //swModel.SketchManager.InsertSketch(false);
            //swModel.SketchManager.AddToDB = true;
            //swModel.SketchManager.DisplayWhenAdded = false;
            //foreach (Circle c in board.circles)
            //{
            //    swModel.SketchManager.CreateCircleByRadius(c.xc, c.yc, 0, c.radius);
            //}
            //swModel.FeatureManager.FeatureCut3(true, false, true, 1, 0, board.thickness, board.thickness, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, true, true, false, 0, 0, false);

            //swAssy.HideComponent();
            //swAssy.ShowComponent();

            //foreach (Point p in board.point)
            //{
            //    Console.WriteLine(p.x);
            //    Console.WriteLine(p.y);
            //    swModel.SketchManager.CreatePoint(p.x, p.y, 0);
            //}

            //swModel.ClearSelection2(true);
            //swAssy.EditAssembly();

            //for (int i = 0; i < board_outline.Count; i++) { Console.WriteLine(board_outline[i]); }
            //for (int i = 0; i < drilled_holes.Count; i++) { Console.WriteLine(drilled_holes[i]); }
            //for (int i = 0; i < placement.Count; i++) { Console.WriteLine(placement[i]); }

            //Console.ReadKey();
            return new XDocument();
        }
    }
}
