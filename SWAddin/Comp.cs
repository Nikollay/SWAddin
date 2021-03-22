using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Threading;

namespace SWAddin
{
    class Comp
    {
        private const double eps = 0.0000001;
        public string used, format, designation, name, note, chapter, included, doc = "", type = "", rotation;
        public int quantity;
        public double x, y, z;

        private static string Euler(double[] R)
        {

            double alpha, alpha2, beta, beta2, gamma, gamma2;

            if (Math.Abs(Math.Abs(R[6]) - 1) > eps)
            {
                beta = -Math.Asin(R[6]);
                beta2 = Math.PI - beta;
#pragma warning disable IDE0059 // Ненужное присваивание значения
                alpha = Math.Atan2(R[7] / Math.Cos(beta), R[8] / Math.Cos(beta));
                alpha2 = Math.Atan2(R[7] / Math.Cos(beta2), R[8] / Math.Cos(beta2));
                gamma = Math.Atan2(R[3] / Math.Cos(beta), R[0] / Math.Cos(beta));
                gamma2 = Math.Atan2(R[3] / Math.Cos(beta2), R[0] / Math.Cos(beta2));
#pragma warning restore IDE0059 // Ненужное присваивание значения
            }
            else
            {
                gamma = 0;
            }
            if (Math.Abs(R[6] + 1) < eps)
            {
                beta = Math.PI / 2;
                alpha = gamma + Math.Atan2(R[1], R[2]);
            }
            else
            {
                beta = -Math.PI / 2;
                alpha = -gamma + Math.Atan2(-R[1], -R[2]);
            }

            return Math.Round(alpha * 180 / Math.PI, 2) + "; " + Math.Round(beta * 180 / Math.PI, 2) + "; " + Math.Round(gamma * 180 / Math.PI, 2);
        }

        public static List<Comp> GetColl(SldWorks swApp)
        {
            AssemblyDoc swAssy;
            Comp component;
            List<Comp> coll;
            object[] comps;
            Component2 comp;
            ModelDoc2 compDoc, swModel;
            CustomPropertyManager prpMgr;
            ModelDocExtension swModelDocExt;
            swDocumentTypes_e docType = swDocumentTypes_e.swDocPART;
            ConfigurationManager confManager;
            string configuration;
            double[] aTrans;
            string path;

            coll = new List<Comp>();
            //swModel.ShowConfiguration2(f.conf[i]);
            swAssy = (AssemblyDoc)swApp.ActiveDoc;
            swAssy.ResolveAllLightWeightComponents(false);
            comps = (object[])swAssy.GetComponents(true);
            
            
            for (int i = 0; i < comps.Length; i++)
            {
                           
                    component = new Comp();
                    swModel = (ModelDoc2)swAssy;
                    swModelDocExt = swModel.Extension;


                    confManager = (ConfigurationManager)swModel.ConfigurationManager;
                    configuration = confManager.ActiveConfiguration.Name;
                    prpMgr = swModelDocExt.get_CustomPropertyManager(configuration);
                    prpMgr.Get6("Обозначение", true, out string valOut, out _, out _, out _);
                    component.used = valOut;
               
                    comp = (Component2)comps[i];
                //проверка компонента
                    if (comp == null)
                    {
                        swApp.SendMsgToUser2("Не могу загрузить " + comp.Name2, 4, 2);
                        continue;
                    }
                    path = comp.GetPathName();
                    if ((comp.GetSuppression() != (int)swComponentSuppressionState_e.swComponentSuppressed) & (comp != null))//comps[i] != null
                    {

                        aTrans = (double[])comp.Transform2.ArrayData;
                        if (path.ToUpper().EndsWith(".SLDASM")) { docType = (swDocumentTypes_e)swDocumentTypes_e.swDocASSEMBLY; }
                        if (path.ToUpper().EndsWith(".SLDPRT")) { docType = (swDocumentTypes_e)swDocumentTypes_e.swDocPART; }
                        int errs = 0, wrns = 0;
                        compDoc = swApp.OpenDoc6(path, (int)docType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errs, ref wrns);
                        if (compDoc == null) { compDoc = (ModelDoc2)comp.GetModelDoc2(); }
                        if (compDoc == null)
                        {
                            swApp.SendMsgToUser2("Не могу загрузить " + path, 4, 2);
                            swApp.ExitApp();
                            System.Environment.Exit(0);
                        }
                        configuration = (string)comp.ReferencedConfiguration;
                        swModelDocExt = (ModelDocExtension)compDoc.Extension;
                        prpMgr = (CustomPropertyManager)swModelDocExt.get_CustomPropertyManager(configuration);

                        prpMgr.Get6("Формат", true, out valOut, out _, out _, out _);
                        component.format = valOut;
                        prpMgr.Get6("Обозначение", true, out valOut, out _, out _, out _);
                        component.designation = valOut;
                        prpMgr.Get6("Наименование", true, out valOut, out _, out _, out _);
                        component.name = valOut;
                        prpMgr.Get6("Примечание", true, out valOut, out _, out _, out _);
                        component.note = valOut;
                        prpMgr.Get6("Раздел", true, out valOut, out _, out _, out _);
                        component.chapter = valOut;
                        prpMgr.Get6("Перв.Примен.", true, out valOut, out _, out _, out _);
                        component.included = valOut;

                        if ((component.chapter == "Стандартные изделия") | (component.chapter == "Прочие изделия"))
                        {
                            prpMgr.Get6("Документ на поставку", true, out valOut, out _, out _, out _);
                            component.doc = valOut;
                            component.type = component.name.Substring(0, component.name.IndexOf((char)32));
                        }

                        component.x = Math.Round((aTrans[9] * 1000), 2);
                        component.y = Math.Round((aTrans[10] * 1000), 2);
                        component.z = Math.Round((aTrans[11] * 1000), 2);
                        component.rotation = Euler(aTrans);
                        component.quantity = 1;

                        coll.Add(component);
                    }
                
            }

            foreach (Comp k in coll)
            {
                if (k.chapter != "Сборочные единицы" & k.chapter != "Детали" & k.chapter != "Документация" & k.chapter != "Комплекты")
                {
                    k.format = "";
                }

                if (k.chapter != "Сборочные единицы" & k.chapter != "Детали" & k.chapter != "Документация" & k.chapter != "Комплекты" & k.chapter != "Стандартные изделия")
                {
                   k.designation = "";
                }

                if (k.format.Contains("*)"))
                {
                    k.note = k.format.Substring(2);
                    k.format = "*)";
                }
            }

            return coll;
        }
        public static XElement GetComponent(Comp comp)
        {

            XAttribute[] name, value;
            name = new XAttribute[29];
            value = new XAttribute[29];
            XElement property, properties, component;
            properties = new XElement("properties");
            component = new XElement("component");

            name[0] = new XAttribute("name", "Раздел СП");
            value[0] = new XAttribute("value", comp.chapter);
            name[1] = new XAttribute("name", "Раздел ВП");
            value[1] = new XAttribute("value", "");
            name[2] = new XAttribute("name", "Количество");
            value[2] = new XAttribute("value", "");
            name[3] = new XAttribute("name", "Подраздел СП");
            value[3] = new XAttribute("value", comp.type);
            name[4] = new XAttribute("name", "Подраздел ВП");
            value[4] = new XAttribute("value", "");
            name[5] = new XAttribute("name", "Примечание");
            value[5] = new XAttribute("value", comp.note);
            name[6] = new XAttribute("name", "Количество на изд.");
            value[6] = new XAttribute("value", "");
            name[7] = new XAttribute("name", "Количество в комп.");
            value[7] = new XAttribute("value", "");
            name[8] = new XAttribute("name", "Количество на рег.");
            value[8] = new XAttribute("value", "");
            name[9] = new XAttribute("name", "Наличие компонента");
            value[9] = new XAttribute("value", "1");
            name[10] = new XAttribute("name", "Позиционное обозначение");
            value[10] = new XAttribute("value", "");
            name[11] = new XAttribute("name", "Наименование");
            value[11] = new XAttribute("value", comp.name);
            name[12] = new XAttribute("name", "Обозначение");
            value[12] = new XAttribute("value", comp.designation);
            name[13] = new XAttribute("name", "Формат");
            value[13] = new XAttribute("value", comp.format);
            name[14] = new XAttribute("name", "Зона");
            value[14] = new XAttribute("value", "");
            name[15] = new XAttribute("name", "Позиция");
            value[15] = new XAttribute("value", "");
            name[16] = new XAttribute("name", "Документ на поставку");
            value[16] = new XAttribute("value", comp.doc);
            name[17] = new XAttribute("name", "Код продукции");
            value[17] = new XAttribute("value", "");
            name[18] = new XAttribute("name", "Поставщик");
            value[18] = new XAttribute("value", "");
            name[19] = new XAttribute("name", "Производитель");
            value[19] = new XAttribute("value", "");
            name[20] = new XAttribute("name", "Тип");
            value[20] = new XAttribute("value", comp.type);
            name[21] = new XAttribute("name", "Куда входит");
            value[21] = new XAttribute("value", comp.used);
            name[22] = new XAttribute("name", "Footprint");
            value[22] = new XAttribute("value", "");
            name[23] = new XAttribute("name", "X");
            value[23] = new XAttribute("value", comp.x);
            name[24] = new XAttribute("name", "Y");
            value[24] = new XAttribute("value", comp.y);
            name[25] = new XAttribute("name", "Z");
            value[25] = new XAttribute("value", comp.z);
            name[26] = new XAttribute("name", "Layer");
            value[26] = new XAttribute("value", "1");
            name[27] = new XAttribute("name", "Rotation");
            value[27] = new XAttribute("value", comp.rotation);
            name[28] = new XAttribute("name", "StandOff");
            value[28] = new XAttribute("value", "0");

            for (int i = 0; i < name.Length; i++)
            {
                property = new XElement("property", name[i], value[i]);
                properties.Add(property);
            }
            component.Add(properties);
            //Console.WriteLine(component);
            return component;
        }
        public static XElement GetGraphs(AssemblyDoc swAssy)
        {
            XAttribute[] name, value;
            name = new XAttribute[27];
            value = new XAttribute[27];
            XElement graph, graphs;
            graphs = new XElement("graphs");

            string approved, developed, project, verified, note, included, designation, title, normal_inspection;

            ModelDoc2 swModel;
            CustomPropertyManager prpMgr;
            ModelDocExtension swModelDocExt;
            ConfigurationManager confManager;
            string configuration;

            swModel = (ModelDoc2)swAssy;

            confManager = (ConfigurationManager)swModel.ConfigurationManager;
            configuration = confManager.ActiveConfiguration.Name;
            swModelDocExt = swModel.Extension;
            prpMgr = swModelDocExt.get_CustomPropertyManager(configuration);
            prpMgr.Get6("п_Утв", true, out string valOut, out _, out _, out _);
            approved = valOut;
            prpMgr.Get6("п_Разраб", true, out valOut, out _, out _, out _);
            developed = valOut;
            prpMgr.Get6("Проект", true, out valOut, out _, out _, out _);
            project = valOut;
            prpMgr.Get6("п_Пров", true, out valOut, out _, out _, out _);
            verified = valOut;
            prpMgr.Get6("Примечание", true, out valOut, out _, out _, out _);
            note = valOut;
            prpMgr.Get6("Перв. примен.", true, out valOut, out _, out _, out _);
            included = valOut;
            prpMgr.Get6("Обозначение", true, out valOut, out _, out _, out _);
            designation = valOut;
            prpMgr.Get6("Наименование", true, out valOut, out _, out _, out _);
            title = valOut;
            prpMgr.Get6("п_Н_контр", true, out valOut, out _, out _, out _);
            normal_inspection = valOut;

            name[0] = new XAttribute("name", "Шифр PCB");
            value[0] = new XAttribute("value", "");
            name[1] = new XAttribute("name", "Характер работы");
            value[1] = new XAttribute("value", "");
            name[2] = new XAttribute("name", "Формат PCB");
            value[2] = new XAttribute("value", "");
            name[3] = new XAttribute("name", "Утвердил");
            value[3] = new XAttribute("value", approved);
            name[4] = new XAttribute("name", "Указания изменение");
            value[4] = new XAttribute("value", "");
            name[5] = new XAttribute("name", "Разработал схемотехник");
            value[5] = new XAttribute("value", "");
            name[6] = new XAttribute("name", "Разработал конструктор");
            value[6] = new XAttribute("value", developed);
            name[7] = new XAttribute("name", "Раздел");
            value[7] = new XAttribute("value", "Документация");
            name[8] = new XAttribute("name", "Проект");
            value[8] = new XAttribute("value", project);
            name[9] = new XAttribute("name", "Проверил схемотехник");
            value[9] = new XAttribute("value", "");
            name[10] = new XAttribute("name", "Проверил конструктор");
            value[10] = new XAttribute("value", verified);
            name[11] = new XAttribute("name", "Примечание");
            value[11] = new XAttribute("value", note);
            name[12] = new XAttribute("name", "Порядковый номер изменения");
            value[12] = new XAttribute("value", "");
            name[13] = new XAttribute("name", "Первичная применяемость");
            value[13] = new XAttribute("value", included);
            name[14] = new XAttribute("name", "Обозначение PCB");
            value[14] = new XAttribute("value", "");
            name[15] = new XAttribute("name", "Обозначение");
            value[15] = new XAttribute("value", designation);
            name[16] = new XAttribute("name", "Нормоконтроль");
            value[16] = new XAttribute("value", normal_inspection);
            name[17] = new XAttribute("name", "Номер документа изменение");
            value[17] = new XAttribute("value", "");
            name[18] = new XAttribute("name", "Наименование PCB");
            value[18] = new XAttribute("value", "");
            name[19] = new XAttribute("name", "Наименование");
            value[19] = new XAttribute("value", title);
            name[20] = new XAttribute("name", "Литера 1");
            value[20] = new XAttribute("value", "");
            name[21] = new XAttribute("name", "Литера 2");
            value[21] = new XAttribute("value", "");
            name[22] = new XAttribute("name", "Литера 3");
            value[22] = new XAttribute("value", "");
            name[23] = new XAttribute("name", "Код документа");
            value[23] = new XAttribute("value", "");
            name[24] = new XAttribute("name", "Дополнительная графа");
            value[24] = new XAttribute("value", "");
            name[25] = new XAttribute("name", "Дата изменения");
            value[25] = new XAttribute("value", "");
            name[26] = new XAttribute("name", "Вид документа");
            value[26] = new XAttribute("value", "Электронная модель сборочной единицы");

            for (int i = 0; i < name.Length; i++)
            {
                graph = new XElement("graph", name[i], value[i]);
                graphs.Add(graph);
            }
            return graphs;
        }
        public static XElement GetDocuments(AssemblyDoc swAssy)
        {
            XAttribute[] name, value;
            name = new XAttribute[15];
            value = new XAttribute[15];
            XElement property, properties, document, documents;
            properties = new XElement("properties");
            document = new XElement("document");
            documents = new XElement("documents");

            string title;

            ModelDoc2 swModel;
            CustomPropertyManager prpMgr;
            ModelDocExtension swModelDocExt;
            ConfigurationManager confManager;
            string configuration;

            swModel = (ModelDoc2)swAssy;

            confManager = (ConfigurationManager)swModel.ConfigurationManager;
            configuration = confManager.ActiveConfiguration.Name;
            swModelDocExt = swModel.Extension;
            prpMgr = swModelDocExt.get_CustomPropertyManager(configuration);
            prpMgr.Get6("Обозначение", true, out string valOut, out _, out _, out _);
            title = valOut;
            

            name[0] = new XAttribute("name", "Раздел СП");
            value[0] = new XAttribute("value", "Документация");
            name[1] = new XAttribute("name", "Наименование");
            value[1] = new XAttribute("value", "Сборочный чертеж");
            name[2] = new XAttribute("name", "Обозначение");
            value[2] = new XAttribute("value", title.Substring(0,15) + "СБ");
            name[3] = new XAttribute("name", "Код продукции");
            value[3] = new XAttribute("value", "");
            name[4] = new XAttribute("name", "Формат");
            value[4] = new XAttribute("value", "A3");
            name[5] = new XAttribute("name", "Документ на поставку");
            value[5] = new XAttribute("value", "");
            name[6] = new XAttribute("name", "Поставщик");
            value[6] = new XAttribute("value", "");
            name[7] = new XAttribute("name", "Количество на изд.");
            value[7] = new XAttribute("value", "1");
            name[8] = new XAttribute("name", "Количество в комп.");
            value[8] = new XAttribute("value", "");
            name[9] = new XAttribute("name", "Количество на рег.к");
            value[9] = new XAttribute("value", "");
            name[10] = new XAttribute("name", "Количество");
            value[10] = new XAttribute("value", "");
            name[11] = new XAttribute("name", "Примечание");
            value[11] = new XAttribute("value", "");
            name[12] = new XAttribute("name", "Зона");
            value[12] = new XAttribute("value", "");
            name[13] = new XAttribute("name", "Позиция");
            value[13] = new XAttribute("value", "");
            name[14] = new XAttribute("name", "Куда входит");
            value[14] = new XAttribute("value", title);
           
            for (int i = 0; i < name.Length; i++)
            {
                property = new XElement("property", name[i], value[i]);
                properties.Add(property);
            }
            document.Add(properties);
            documents.Add(document);
            return documents;
        }
    }
}
