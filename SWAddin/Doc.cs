using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace drw_tif
{
    class Doc
    {
        public static void GetTiff()
        {
            var progId = "SldWorks.Application.27";

            var progType = System.Type.GetTypeFromProgID(progId);

            var swApp = System.Activator.CreateInstance(progType) as ISldWorks;
            swApp.Visible = false;
            Console.WriteLine(swApp.RevisionNumber());
            Console.CursorSize = 100;
            ModelDoc2 swModel;
            ModelDocExtension swModelDocExt;
            AssemblyDoc swAssy;
            Component2 swComp;
            DrawingDoc Part;

            int errors = 0;
            int warnings = 0;
            string fileName;   // GetOpenFileName
            Dictionary<string, string> Dict, Drw;
            string projekt_path, key, pathName;
            string[] сonfNames;
            object[] Comps;

            fileName = swApp.GetOpenFileName("File to SLDASM", "", "SLDASM Files (*.SLDASM)|*.SLDASM|", out _, out _, out _);
            //Проверяем путь
            if (fileName == "")
            {
                swApp.ExitApp();
                return;
            }
            swModel = (ModelDoc2)swApp.OpenDoc6(fileName, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

            //Проверяем открыта сборка или нет
            if ((swModel.GetType() != 2) | (swModel == null))
            {
                swApp.SendMsgToUser2("Откройте сборку", 4, 2);
                swApp.ExitApp();
                return;
            }

            //swModel = (ModelDoc2)swApp.ActiveDoc;
            //Console.WriteLine(swModel.GetType());

            swAssy = (AssemblyDoc)swModel;
            Dict = new Dictionary<string, string>();
            projekt_path = swModel.GetPathName().Remove(swModel.GetPathName().LastIndexOf((char)92) + 1);

            Console.WriteLine(projekt_path);
            key = swModel.GetPathName().Substring(swModel.GetPathName().LastIndexOf((char)92) + 1);
            key = key.Substring(0, key.Length - 7);
            pathName = swModel.GetPathName();
            pathName = pathName.Remove(pathName.Length - 7);
            Console.WriteLine(key);
            Dict.Add(key, pathName);

            //Создаем список путей компонентов для всех конфигураций
            сonfNames = (string[])swModel.GetConfigurationNames();
            swAssy.ResolveAllLightWeightComponents(false);
            for (int i = 0; i < сonfNames.Length; i++)
            {
                Console.WriteLine((string)сonfNames[i]);
                swModel.ShowConfiguration2((string)сonfNames[i]);
                swModel.ForceRebuild3(false);
                Comps = (Object[])swAssy.GetComponents(false);
                for (int j = 0; j < Comps.Length; j++)
                {
                    swComp = (Component2)Comps[j];
                    //compDoc = (ModelDoc2)swComp.GetModelDoc2();
                    //if (swComp == null) { Console.WriteLine(swComp.Name2); }
                    if ((swComp.GetSuppression() != (int)swComponentSuppressionState_e.swComponentSuppressed) & (swComp != null))
                    {
                        pathName = swComp.GetPathName();
                        Console.WriteLine(pathName);
                        pathName = pathName.Remove(pathName.Length - 7);
                        key = swComp.GetPathName().Substring(swComp.GetPathName().LastIndexOf((char)92) + 1);
                        key = key.Substring(0, key.Length - 7);
                        if (!Dict.ContainsKey(key)) { Dict.Add(key, pathName); }
                    }
                }
                Console.WriteLine("********************************");
            }
            //Console.ReadKey();
            //Находим где могут быть чертежи
            Drw = new Dictionary<string, string>();
            foreach (KeyValuePair<string, string> k in Dict)
            {
                if ((k.Value.Contains((string)"D:\\PDM\\Проект")) | (k.Value.Contains("D:\\PDM\\Общеприменяемые")))
                {
                    Drw.Add(k.Key, k.Value);
                }
            }

            //Создаем папку
            DirectoryInfo dirInfo = new DirectoryInfo(projekt_path + "\\TIF");
            if (!dirInfo.Exists)
            {
                dirInfo.Create();
            }
            Console.WriteLine("Чертежей не более " + Drw.Count);

            //Настройки TIF
            swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffScreenOrPrintCapture, 1); //1-Print capture
            swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffPrintPaperSize, 12); //12-Papers User Defined
            swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffImageType, 0); //0-Black And White
            swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffCompressionScheme, 2); //2-Group 4 Fax Compression
            swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffPrintDPI, 600); //300-Integer value

            //Сохраняем картинки
            int itogo = 0;
            foreach (KeyValuePair<string, string> k in Drw)
            {
                //Настройка размеров картинки
                swApp.IGetTemplateSizes(k.Value + ".SLDDRW", out int PaperSize, out double Width, out double Height);
                swApp.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperWidth, Width); //Double value in meters
                swApp.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperHeight, Height); //Double value in meters

                Part = (DrawingDoc)swApp.OpenDoc6(k.Value + ".SLDDRW", (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", ref errors, ref warnings);
                Console.WriteLine(k.Value + ".SLDDRW");
                if ((errors == 0) & (Part != null))
                {
                    swModel = (ModelDoc2)Part;
                    swModelDocExt = (ModelDocExtension)swModel.Extension;
                    swModelDocExt.SaveAs(projekt_path + "TIF\\" + k.Key + ".TIF", 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errors, ref warnings);
                    itogo += 1;
                }
                swApp.CloseDoc(k.Value + ".SLDDRW");
                Part = null;
            }
            // Console.ReadKey();
            swApp.SendMsgToUser2("Всего частей " + Dict.Count + System.Environment.NewLine + "Чертежей сохранено " + itogo, 2, 2);
            swApp.ExitApp();
            //swApp = null;
        }
    }
}
