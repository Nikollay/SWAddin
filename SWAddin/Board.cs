using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Windows.Forms;
using System.Linq;

using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Globalization;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace SWAddin
{
    class Board
    {
        public List<Component> components;
        public List<Circle> circles;
        //public List<Point> point;
        public List<object> sketh, cutout;
        public double thickness;
        public int ver;
        public string designator;

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
            try
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
            catch
            {
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
                tmpXEl = elements.First(item => item.Attribute("name").Value.Equals("Обозначение_PCB")|item.Attribute("name").Value.Equals("Обозначение PCB"));
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
                        component.x = double.Parse(e.Attribute("value").Value.Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                        break;
                    case "Y":
                        component.y = double.Parse(e.Attribute("value").Value.Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                        break;
                    case "Z":
                        component.z = double.Parse(e.Attribute("value").Value.Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                        break;
                    case "Layer":
                        component.layer = int.Parse(e.Attribute("value").Value.Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture);
                        break;
                    case "Rotation":
                        component.rotation = double.Parse(e.Attribute("value").Value.Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture);
                        break;
                    case "StandOff":
                        component.standOff = double.Parse(e.Attribute("value").Value.Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
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
            board.designator = el.FirstAttribute.Value;
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
                                skLine.x1 = double.Parse(strToDbl[0].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skLine.y1 = double.Parse(strToDbl[1].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skLine.x2 = double.Parse(strToDbl[2].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skLine.y2 = double.Parse(strToDbl[3].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                board.sketh.Add(skLine);
                            }
                            if (strToDbl.Length == 9)
                            {
                                skArc = new Arc();
                                skArc.x1 = double.Parse(strToDbl[0].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skArc.y1 = double.Parse(strToDbl[1].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skArc.x2 = double.Parse(strToDbl[2].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skArc.y2 = double.Parse(strToDbl[3].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skArc.xc = double.Parse(strToDbl[7].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skArc.yc = double.Parse(strToDbl[8].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skArc.a1 = double.Parse(strToDbl[4].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture);
                                skArc.a2 = double.Parse(strToDbl[5].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture);
                                skArc.r = double.Parse(strToDbl[6].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                if (skArc.a1 < skArc.a2) { skArc.direction = 1; }
                                else { skArc.direction = -1; }
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
                                skLine.x1 = double.Parse(strToDbl[0].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skLine.y1 = double.Parse(strToDbl[1].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skLine.x2 = double.Parse(strToDbl[2].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skLine.y2 = double.Parse(strToDbl[3].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                board.cutout.Add(skLine);
                            }
                            if (strToDbl.Length == 9)
                            {
                                skArc = new Arc();
                                skArc.x1 = double.Parse(strToDbl[0].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skArc.y1 = double.Parse(strToDbl[1].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skArc.x2 = double.Parse(strToDbl[2].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skArc.y2 = double.Parse(strToDbl[3].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skArc.xc = double.Parse(strToDbl[7].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skArc.yc = double.Parse(strToDbl[8].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
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
                                skCircle.radius = double.Parse(strToDbl[0].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 2000;
                                skCircle.xc = double.Parse(strToDbl[1].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                skCircle.yc = double.Parse(strToDbl[2].Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
                                board.circles.Add(skCircle);
                            }
                        }
                        break;
                    case "Толщина, мм":
                        board.thickness = double.Parse(e.Attribute("value").Value.Replace(",", CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator), CultureInfo.InvariantCulture) / 1000;
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
            if (string.IsNullOrWhiteSpace(filename)) { System.Environment.Exit(0); }
            return filename;
        }

        public static ExcelPackage GetfromXDocument(XDocument doc, string pdm_path)
        {
            IEnumerable<XElement> elements1, elements2;
            ExcelPackage pck = new OfficeOpenXml.ExcelPackage(new FileInfo(pdm_path + "EPDM_LIBRARY\\EPDM_SolidWorks\\ADDIN\\sp.xlsx"), false);
            ExcelWorksheet wh, wh1;
            ExcelRange wc;
            XElement tmpXEl;
            string designation, type;
            //Провека GostDoc или нет
            //tmpXEl = doc.Root.Element("transaction").Elements().First(item => item.Attribute("Type").Value.Equals("GostDoc"));
            type = doc.Root.Element("transaction").Attribute("Type").Value;

            //Заполняем шапку

            wh = pck.Workbook.Worksheets[1];
            try
            { 
            elements1 = doc.Root.Element("transaction").Element("project").Element("configurations").Element("configuration").Element("graphs").Elements();
            }
            catch
            {
                return null;
            }
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("Проект"));
            wh.Cells[1, 1].Value = tmpXEl.Attribute("value").Value;
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("Первичная применяемость")|item.Attribute("name").Value.Equals("Перв.Примен."));
            wh.Cells[1, 3].Value = tmpXEl.Attribute("value").Value;
            wh.Cells[3, 14].Value = "Документация";
            wc = wh.Cells[3, 14];
            wc.Style.Font.UnderLine = true;
            wc.Style.Font.Bold = true;
            wc.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // xlCenterF
            wc.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // xlCenter
            wh.Cells[5, 4].Value = "A3";
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("Обозначение"));
            designation = tmpXEl.Attribute("value").Value;
            wh.Cells[5, 9].Value = tmpXEl.Attribute("value").Value + "СБ";
            wh.Cells[5, 14].Value = "Сборочный чертеж";
            wh.Cells[32, 12].Value = tmpXEl.Attribute("value").Value;
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Equals("Наименование"));
            wh.Cells[35, 12].Value = tmpXEl.Attribute("value").Value;
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Contains("Разработал конструктор")|item.Attribute("name").Value.Contains("п_Разраб"));
            wh.Cells[35, 8].Value = tmpXEl.Attribute("value").Value;
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Contains("Проверил конструктор")|item.Attribute("name").Value.Contains("п_Пров_P"));
            wh.Cells[36, 8].Value = tmpXEl.Attribute("value").Value;
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Contains("Нормоконтроль")|item.Attribute("name").Value.Contains("п_Н_контр"));
            wh.Cells[38, 8].Value = tmpXEl.Attribute("value").Value;
            tmpXEl = elements1.First(item => item.Attribute("name").Value.Contains("Утвердил")|item.Attribute("name").Value.Contains("п_Утв"));
            wh.Cells[39, 8].Value = tmpXEl.Attribute("value").Value;

            //Заполняем словарь
            try
            { 
            elements1 = doc.Root.Element("transaction").Element("project").Element("configurations").Element("configuration").Element("components").Elements();
            }
            catch
            {
                return null;
            }

            Record component;
            SortedDictionary<string, Record> dictS;
            dictS = new SortedDictionary<string, Record>();
            List<Record> list;
            list = new List<Record>();
            string key;

            if (type == "GostDoc")
            {
                foreach (XElement e1 in elements1)
                {
                    component = new Record();
                    elements2 = e1.Element("properties").Elements();
                    tmpXEl = elements2.First(item => item.Attribute("name").Value.Equals("Наименование"));
                    if (String.IsNullOrEmpty(tmpXEl.Attribute("value").Value))
                    {
                        component.format = " ";
                        component.designation = " ";
                        component.title = " ";
                        component.note = " ";
                        component.chapter = elements2.First(item => item.Attribute("name").Value.Equals("Раздел СП")).Attribute("value").Value;
                        component.pos = " ";
                        component.count = " ";
                    }
                    else
                        {
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
                                case "Позиция":
                                    component.pos = e2.Attribute("value").Value;
                                    break;
                                case "Количество":
                                    component.count = e2.Attribute("value").Value;
                                    break;
                                }
                            }
                    }
                list.Add(component);
                }
            }
            
            else
            {
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
                            case "Позиция":
                                component.pos = e2.Attribute("value").Value;
                                break;
                            case "Количество":
                                component.count = e2.Attribute("value").Value;
                                break;
                        }
                    }
                    key = component.designation + (char)32 + component.title;
                    if (!dictS.ContainsKey(key)) { dictS.Add(key, component); }
                    else dictS[key].quantity++;
                }
                //Заполнили словарь *******
                //Сортировка
                var dict = dictS.GroupBy(g => g.Value.chapter).OrderBy(n => n.Key, new CustomComparer()).ToDictionary(group => group.Key, group => group.ToDictionary(pair => pair.Key, pair => pair.Value));
                foreach (var d in dict.Values)
                {
                    foreach (var v in d.Values) { list.Add(v); }
                }
            }
            string partition = "Документация";
            int j = 6;
            //MessageBox.Show(list.Count.ToString());
            //Заполняем листы

            if (type == "GostDoc")
            {
                foreach (Record lr in list)
                {
                    if (!lr.chapter.Equals(partition))
                    {
                        wc = wh.Cells[j + 2, 14];
                        wh.Cells[j + 2, 14].Value = lr.chapter;
                        wc.Style.Font.UnderLine = true;
                        wc.Style.Font.Bold = true;
                        wc.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // xlCenterF
                        wc.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // xlCenter
                        j += 5;
                        partition = lr.chapter;
                    }

                    if ((j > 26) & (wh.Name.Equals("1")))
                    {
                        wh1 = pck.Workbook.Worksheets[pck.Workbook.Worksheets.Count - 1];
                        ExcelWorksheet excelWorksheet = pck.Workbook.Worksheets.Add("Лист " + (wh1.Index + 1).ToString(), wh1);
                        pck.Workbook.Worksheets.MoveBefore("Лист " + (wh1.Index + 1).ToString(), "ЛРИ");
                        wh = pck.Workbook.Worksheets[pck.Workbook.Worksheets.Count - 2];
                        j = 3;
                    }

                    if (j > 33)
                    {
                        wh1 = pck.Workbook.Worksheets[pck.Workbook.Worksheets.Count - 1];
                        ExcelWorksheet excelWorksheet = pck.Workbook.Worksheets.Add("Лист " + (wh1.Index + 1).ToString(), wh1);
                        pck.Workbook.Worksheets.MoveBefore("Лист " + (wh1.Index + 1).ToString(), "ЛРИ");
                        wh = pck.Workbook.Worksheets[pck.Workbook.Worksheets.Count - 2];
                        j = 3;
                    }

                    wh.Cells[j, 4].Value = lr.format;
                    wh.Cells[j, 7].Value = lr.pos;
                    wh.Cells[j, 9].Value = lr.designation;
                    wh.Cells[j, 20].Value = lr.count;
                    wh.Cells[j, 21].Value = lr.note;

                    if (lr.title.Length < 33) { wh.Cells[j, 14].Value = lr.title; }

                    if (lr.title.Length > 32)
                    {
                        wh.Cells[j, 14].Value = lr.title.Substring(0, 31);
                        wh.Cells[j + 1, 14].Value = lr.title.Substring(31);
                        j += 1;
                    }
                    j += 1;
                }
            }
            else
            {
                foreach (Record lr in list)
                {
                    if ((j % 4) == 0) { j++; }
                    if (!lr.chapter.Equals(partition))
                    {
                        wc = wh.Cells[j + 2, 14];
                        wh.Cells[j + 2, 14].Value = lr.chapter;
                        wc.Style.Font.UnderLine = true;
                        wc.Style.Font.Bold = true;
                        wc.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // xlCenterF
                        wc.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // xlCenter
                        j += 5;
                        partition = lr.chapter;
                    }

                    if ((j > 26) & (wh.Name.Equals("1")))
                    {
                        wh1 = pck.Workbook.Worksheets[pck.Workbook.Worksheets.Count - 1];
                        ExcelWorksheet excelWorksheet = pck.Workbook.Worksheets.Add("Лист " + (wh1.Index + 1).ToString(), wh1);
                        pck.Workbook.Worksheets.MoveBefore("Лист " + (wh1.Index + 1).ToString(), "ЛРИ");
                        wh = pck.Workbook.Worksheets[pck.Workbook.Worksheets.Count - 2];
                        j = 4;
                    }

                    if (j > 33)
                    {
                        wh1 = pck.Workbook.Worksheets[pck.Workbook.Worksheets.Count - 1];
                        ExcelWorksheet excelWorksheet = pck.Workbook.Worksheets.Add("Лист " + (wh1.Index + 1).ToString(), wh1);
                        pck.Workbook.Worksheets.MoveBefore("Лист " + (wh1.Index + 1).ToString(), "ЛРИ");
                        wh = pck.Workbook.Worksheets[pck.Workbook.Worksheets.Count - 2];
                        j = 4;
                    }

                    wh.Cells[j, 4].Value = lr.format;
                    wh.Cells[j, 9].Value = lr.designation;
                    wh.Cells[j, 20].Value = lr.quantity;
                    wh.Cells[j, 21].Value = lr.note;

                    if (lr.title.Length < 33) { wh.Cells[j, 14].Value = lr.title; }

                    if (lr.title.Length > 32)
                    {
                        wh.Cells[j, 14].Value = lr.title.Substring(0, 31);
                        wh.Cells[j + 1, 14].Value = lr.title.Substring(31);
                        j += 1;
                    }
                    j += 1;
                }
            }
            //Заполнили
            
            pck.Workbook.Worksheets.Delete(pck.Workbook.Worksheets.Count - 1);
            //Удаляем лист шаблон

            if (pck.Workbook.Worksheets.Count == 2)
            {
                wh = pck.Workbook.Worksheets[1];
                wh.Cells[36, 19].Value = "";
            }
            if (pck.Workbook.Worksheets.Count < 4) { pck.Workbook.Worksheets.Delete("ЛРИ"); } //Удаляем лист ЛРИ
            wh = pck.Workbook.Worksheets[1];
            wh.Cells[36, 22].Value = pck.Workbook.Worksheets.Count;

            for (int i = 2; i < pck.Workbook.Worksheets.Count+1; i++)
                {
                    wh = pck.Workbook.Worksheets[i];
                    wh.Cells[35, 12].Value = designation;
                    if (!wh.Name.Equals("ЛРИ"))
                    {
                        wh.Name = i.ToString();
                        wh.Cells[37, 22].Value = i;
                    }
                    if (wh.Name.Equals("ЛРИ")) { wh.Cells[37, 19].Value = pck.Workbook.Worksheets.Count; }
                }
            
        return pck;
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
        public class CustomComparer : IComparer<string>
        {
            public int Compare(string x, string y)

            {
                if (Wt(x) < Wt(y))
                    return -1;
                if (Wt(x) > Wt(y))
                    return 1;
                else return 0;

                //"Документация" "Сборочные единицы" "Стандартные изделия" "Прочие изделия" "Материалы" "Комплекты" 
                // do your own comparison however you like; return a negative value
                // to indicate that x < y, a positive value to indicate that x > y,
                // or 0 to indicate that they are equal.
            }
            private int Wt(string arg)

            {
                switch (arg)
                {
                    case "Документация":
                        return 1;
                    case "Сборочные единицы":
                        return 2;
                    case "Детали":
                        return 3;
                    case "Стандартные изделия":
                        return 4;
                    case "Прочие изделия":
                        return 5;
                    case "Материалы":
                        return 6;
                    case "Комплекты":
                        return 7;
                    default:
                        return 8;
                }
            }
        }
    }
}
