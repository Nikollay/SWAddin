﻿using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Reflection;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using SolidWorksTools;
using SolidWorksTools.File;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.ComponentModel;
using System.Xml.Linq;

namespace SWAddin
{

    [Guid("3d605915-bd53-4097-a24f-ce262db3bab0"), ComVisible(true)]
    [SwAddin(Description = "AddinIN description", Title = "AddinIN", LoadAtStartup = true)]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 1;
        public const int mainItemID2 = 2;
        public const int mainItemID3 = 3;
        string sAddinName = "C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS PDM\\PDMSW.dll";
        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = ThisSW as ISldWorks;
            addinID = cookie;

            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            #region Setup the Command Manager
            iCmdMgr = iSwApp.GetCommandManager(cookie);
            AddCommandMgr();
            #endregion

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;

            int cmdIndex1, cmdIndex2, cmdIndex3;
            string Title = "Addin", ToolTip = "Addin";

            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[3] { mainItemID1, mainItemID2, mainItemID3 };

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
                {
                    ignorePrevious = true;
                }
            }

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "Аддин", -1, ignorePrevious, ref cmdGroupErr);

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            cmdIndex1 = cmdGroup.AddCommandItem2("Создать 3D модель платы", -1, "Создать 3D модель платы PCB", "Создать 3D модель платы PCB", -1, "Create3DPCB", "", mainItemID1, menuToolbarOption);
            cmdIndex2 = cmdGroup.AddCommandItem2("Создать XML", -1, "Создать XML из сборки", "Создать XML из сборки", -1, "GetXML", "", mainItemID2, menuToolbarOption);
            cmdIndex3 = cmdGroup.AddCommandItem2("Создать Tiff", -1, "Создать Tiff картинки чертежей", "Создать Tiff картинки чертежей", -1, "GetTiff", "", mainItemID3, menuToolbarOption);

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();
        }

        public void RemoveCommandMgr()
        {
            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {

                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        #endregion

        #region UI Callbacks

        public void Create3DPCB()
        {
            iSwApp.CommandInProgress = true;
            Board board;
            string filename;
            filename = iSwApp.GetOpenFileName("Открыть файл", "", "xml Files (*.xml)|*.xml|", out _, out _, out _); //Board.GetFilename();
            if (String.IsNullOrWhiteSpace(filename)) { return; }
            board = Board.GetfromXML(filename);
            if (board == null) { MessageBox.Show("XML с неверной структурой","Ошибка чтения файла"); return; }
            ModelDoc2 swModel;
            AssemblyDoc swAssy;
            ModelView activeModelView;

            iSwApp.UnloadAddIn(sAddinName);
            
            //Новая сборка платы
            double swSheetWidth = 0, swSheetHeight = 0;
            string boardName;
            int Errors = 0, Warnings = 0;
            swAssy = (AssemblyDoc)iSwApp.NewDocument("D:\\PDM\\EPDM_LIBRARY\\EPDM_SolidWorks\\EPDM_SWR_Templates\\Модуль_печатной_платы.asmdot", (int)swDwgPaperSizes_e.swDwgPaperA2size, swSheetWidth, swSheetHeight);
            swModel = (ModelDoc2)swAssy;
            //Сохранение
            boardName = filename.Remove(filename.Length - 3) + "SLDASM";
            swModel.Extension.SaveAs(boardName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_UpdateInactiveViews, null, ref Errors, ref Warnings);
            //**********

            //Доска
            Component2 board_body;
            PartDoc part;
            ModelDoc2 swCompModel;
            Feature swRefPlaneFeat, plane;
            swAssy.InsertNewVirtualPart(null, out board_body);
            board_body.Select4(false, null, false);
            swAssy.EditPart();
            swCompModel = (ModelDoc2)board_body.GetModelDoc2();
            part = (PartDoc)swCompModel;
            part.SetMaterialPropertyName2("-00", "гост материалы.sldmat", "Rogers 4003C");

            int j = 1;
            do
            {
                swRefPlaneFeat = (Feature)swCompModel.FeatureByPositionReverse(j);
                j++;
            }
            while (swRefPlaneFeat.Name != "Спереди");

            plane = (Feature)board_body.GetCorresponding(swRefPlaneFeat);
            plane.Select2(false, -1);

            swModel.SketchManager.InsertSketch(false);
            swModel.SketchManager.AddToDB = true;

            //Эскизы
            swModel.SketchManager.DisplayWhenAdded = false;
            foreach (Object skt in board.sketh)
            {
                if (skt.GetType().FullName == "SWAddin.Line") { Line sk = (Line)skt; swModel.SketchManager.CreateLine(sk.x1, sk.y1, 0, sk.x2, sk.y2, 0); }
                if (skt.GetType().FullName == "SWAddin.Arc") { Arc sk = (Arc)skt; swModel.SketchManager.CreateArc(sk.xc, sk.yc, 0, sk.x1, sk.y1, 0, sk.x2, sk.y2, 0, sk.direction); }
            }
            swModel.FeatureManager.FeatureExtrusion3(true, false, false, 0, 0, board.thickness, board.thickness, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);
            swModel.ClearSelection2(true);

            if (board.cutout.Count>2)
            {
                plane.Select2(false, -1);
                swModel.SketchManager.InsertSketch(false);
                swModel.SketchManager.AddToDB = true;
                foreach (Object skt in board.cutout)
                {
                    if (skt.GetType().FullName == "SWAddin.Line") { Line sk = (Line)skt; swModel.SketchManager.CreateLine(sk.x1, sk.y1, 0, sk.x2, sk.y2, 0); }
                    if (skt.GetType().FullName == "SWAddin.Arc") { Arc sk = (Arc)skt; swModel.SketchManager.CreateArc(sk.xc, sk.yc, 0, sk.x1, sk.y1, 0, sk.x2, sk.y2, 0, sk.direction); }
                }
                swModel.FeatureManager.FeatureCut4(true, false, true, 1, 0, board.thickness, board.thickness, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, true, true, false, 0, 0, false, false);
            }


            plane.Select2(false, -1);
            swModel.SketchManager.InsertSketch(false);
            swModel.SketchManager.AddToDB = true;
            
            foreach (Circle c in board.circles)
            {
                swModel.SketchManager.CreateCircleByRadius(c.xc, c.yc, 0, c.radius);
            }
            swModel.FeatureManager.FeatureCut4(true, false, true, 1, 0, board.thickness, board.thickness, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, true, true, false, 0, 0, false, false);
            //swModel.FeatureManager.FeatureCut3(true, false, true, 1, 0, board.thickness, board.thickness, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, true, true, false, 0, 0, false);

            swModel.SketchManager.DisplayWhenAdded = true;
            swModel.SketchManager.AddToDB = false;
            swAssy.HideComponent();
            swAssy.ShowComponent();
            swModel.ClearSelection2(true);
            swAssy.EditAssembly();

            string path, sample;
            switch (board.ver)
            {
                case 1:
                    path = "D:\\PDM\\Прочие изделия\\ЭРИ";
                    break;
                case 2:
                    path = "D:\\PDM\\Прочие изделия\\Footprint";
                    break;
                default:
                    path = "D:\\PDM\\Прочие изделия\\ЭРИ";
                    break;
            }
            List<string> allFoundFiles = new List<string>(Directory.GetFiles(path, "*.SLD*", SearchOption.AllDirectories));
            Dictionary<string, string> empty = new Dictionary<string, string>();
                        
            foreach (Component comp in board.components)
            {
                sample = comp.title;
                if (board.ver==2) { sample = comp.footprint; }
                comp.fileName = allFoundFiles.Find(item => item.Contains(sample));
                if (String.IsNullOrWhiteSpace(comp.fileName))
                {
                    comp.fileName = "D:\\PDM\\Прочие изделия\\ЭРИ\\Zero.SLDPRT";
                    if (!empty.ContainsKey(sample)) { empty.Add(sample, sample); }
                }
            }

            double[] transforms, dMatrix;
            string[] coordSys, names;
            double alfa, beta, gamma, x, y, z;
            names = new string[board.components.Count];
            coordSys = new string[board.components.Count];
            dMatrix = new double[16];
            transforms = new double[board.components.Count * 16];

            for (int i = 0; i < board.components.Count; i++)
            {
                names[i] = board.components[i].fileName;
            }
            int n = 0;
            foreach (Component comp in board.components)
            {
                alfa = 0;
                x = comp.x;
                y = comp.y;
                z = comp.z;
                if (comp.layer == 1) //Если Top
                {
                    //z = (comp.z + comp.standOff) standOff не учитывается
                    beta = -Math.PI / 2;
                }
                else             //Иначе Bottom
                {
                    // z = (comp.z - comp.standOff) standOff не учитывается
                    beta = Math.PI / 2;
                }
                gamma = -(comp.rotation / 180) * Math.PI;

                dMatrix[0] = Math.Cos(alfa) * Math.Cos(gamma) - Math.Sin(alfa) * Math.Cos(beta) * Math.Sin(gamma);
                dMatrix[1] = -Math.Cos(alfa) * Math.Sin(gamma) - Math.Sin(alfa) * Math.Cos(beta) * Math.Cos(gamma);
                dMatrix[2] = Math.Sin(alfa) * Math.Sin(beta); //1 строка матрицы вращения
                dMatrix[3] = Math.Sin(alfa) * Math.Cos(gamma) + Math.Cos(alfa) * Math.Cos(beta) * Math.Sin(gamma);
                dMatrix[4] = -Math.Sin(alfa) * Math.Sin(gamma) + Math.Cos(alfa) * Math.Cos(beta) * Math.Cos(gamma);
                dMatrix[5] = -Math.Cos(alfa) * Math.Sin(beta); //2 строка матрицы вращения
                dMatrix[6] = Math.Sin(beta) * Math.Sin(gamma);
                dMatrix[7] = Math.Sin(beta) * Math.Cos(gamma);
                dMatrix[8] = Math.Cos(beta); //3 строка матрицы вращения
                dMatrix[9] = x; dMatrix[10] = y; dMatrix[11] = z; //Координаты
                dMatrix[12] = 1; //Масштаб
                dMatrix[13] = 0; dMatrix[14] = 0; dMatrix[15] = 0; //Ничего

                for (int k = 0; k < dMatrix.Length; k++) { transforms[n * 16 + k] = dMatrix[k]; }
                n++;
            }

            //Вставка
            swAssy.AddComponents3(names, transforms, coordSys);

            //Фиксация
            swModel.Extension.SelectAll();
            swAssy.FixComponent();
            swModel.ClearSelection2(true);

            activeModelView = (ModelView)swModel.ActiveView;
            activeModelView.DisplayMode = (int)swViewDisplayMode_e.swViewDisplayMode_ShadedWithEdges;
            //****************************

            UserProgressBar pb;
            iSwApp.GetUserProgressBar(out pb);

            //Заполнение поз. обозначений
            List<Component2> compsColl = new List<Component2>(); //Коллекция из компонентов сборки платы
            Feature swFeat;
            Component2 compTemp;
            pb.Start(0, board.components.Count, "Поиск");
            int itm = 0;
            swFeat = (Feature)swModel.FirstFeature();
            while (swFeat != null)
            {
                pb.UpdateProgress(itm);
                //pb.UpdateTitle(itm);
                if (swFeat.GetTypeName().Equals("Reference")) //Заполняем коллекцию изделиями
                {
                    compTemp = (Component2)swFeat.GetSpecificFeature2();
                    compsColl.Add(compTemp);
                }
                swFeat = (Feature)swFeat.GetNextFeature();
                itm++;
            }
            pb.End();

            compsColl[0].Name2 = "Плата"; //Пререименовываем деталь      
            if (compsColl.Count - 1 == board.components.Count) //Проверка чтобы не сбились поз. обозначения, если появятся значит все правильно иначе они не нужны
            {
                for (int i = 0; i < board.components.Count; i++)
                    compsColl[i + 1].ComponentReference = board.components[i].physicalDesignator; //Заполняем поз. обозначениями
            }

            string estr = "";
            
            iSwApp.LoadAddIn(sAddinName);

            if (empty.Count != 0)
            {
                foreach (KeyValuePair<string, string> str in empty) { estr = estr + str.Value + System.Environment.NewLine; }
                MessageBox.Show(estr, "Не найдены");
                //swApp.SendMsgToUser2("Не найдены" + estr, 2, 2);
            }
            iSwApp.CommandInProgress = false;
            //**************

        }
        public void GetXML()
        {
            iSwApp.CommandInProgress = true;
            ModelDoc2 swModel;
            AssemblyDoc swAssy;
            List<Comp> coll;
            XDocument doc;
            XElement xml, transaction, project, configurations, configuration, components, component;

            int errors = 0;
            int warnings = 0;
            string fileName;   // GetOpenFileName
            string path;
            List<string> conf;
            string[] сonfNames;
            swModel = (ModelDoc2)iSwApp.ActiveDoc;
            if (swModel == null)
            {
                fileName = iSwApp.GetOpenFileName("Выберите сборку", "", "SLDASM Files (*.SLDASM)|*.SLDASM|", out _, out _, out _);
                //Проверяем путь
                if (fileName == "")
                {
                    return;
                }
                swModel = (ModelDoc2)iSwApp.OpenDoc6(fileName, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            }

            iSwApp.UnloadAddIn(sAddinName);
            
            fileName = swModel.GetPathName();
            //Проверяем открыта сборка или нет
            if ((swModel.GetType() != 2) | (swModel == null))
            {
                iSwApp.SendMsgToUser2("Откройте сборку", 4, 2);
                return;
            }
            swAssy = (AssemblyDoc)swModel;

            doc = new XDocument(new XDeclaration("1.0", "Windows-1251", "Yes"));
            xml = new XElement("xml");
            transaction = new XElement("transaction", new XAttribute("Type", "SOLIDWORKS"), new XAttribute("version", "1.0"), new XAttribute("Date", DateTime.Now.ToString("d")), new XAttribute("Time", DateTime.Now.ToString("T")));
            project = new XElement("project", new XAttribute("Project_Path", fileName), new XAttribute("Project_Name", swModel.GetTitle() + ".SldAsm"));
            configurations = new XElement("configurations");
            components = new XElement("components");
            сonfNames = (string[])swModel.GetConfigurationNames();
            conf = new List<string>(сonfNames);

            ConfigForm f = new ConfigForm(conf);
            f.ShowDialog();
            if (f.conf == null)
            {
                return;
            }

            if (f.conf.Count == 0)
            {
                return;
            }

            for (int i = 0; i < f.conf.Count; i++)
            {
                swModel.ShowConfiguration2(f.conf[i]);
                configuration = new XElement("configuration", new XAttribute("name", f.conf[i]));
                coll = Comp.GetColl(swAssy, (SldWorks)iSwApp);
                foreach (Comp k in coll)
                {
                    component = Comp.GetComponent(k);
                    components.Add(component);
                }
                if (i == 0) { configuration.Add(Comp.GetGraphs(swAssy)); }
                configuration.Add(components);
                configurations.Add(configuration);
            }
            project.Add(configurations);
            transaction.Add(project);
            xml.Add(transaction);
            doc.Add(xml);
            path = fileName.Substring(0, fileName.Length - 7) + ".xml";
            iSwApp.LoadAddIn(sAddinName);
            doc.Save(path);
            iSwApp.CommandInProgress = false;

        }
        public void GetTiff()
        {
            //object obt= iSwApp.GetAddInObject("ConisioSW2.ConisioSWAddIn") as SwAddin;
            iSwApp.CommandInProgress = true;
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

            swModel = (ModelDoc2)iSwApp.ActiveDoc;
            if (swModel == null)
            {
                fileName = iSwApp.GetOpenFileName("File to SLDASM", "", "SLDASM Files (*.SLDASM)|*.SLDASM|", out _, out _, out _);
                //Проверяем путь
                if (fileName == "")
                {

                    return;
                }
                swModel = (ModelDoc2)iSwApp.OpenDoc6(fileName, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            }
            iSwApp.UnloadAddIn(sAddinName);
            //Проверяем открыта сборка или нет
            if ((swModel.GetType() != 2) | (swModel == null))
            {
                iSwApp.SendMsgToUser2("Откройте сборку", 4, 2);
                
                return;
            }

            //swModel = (ModelDoc2)swApp.ActiveDoc;
            
            swAssy = (AssemblyDoc)swModel;
            Dict = new Dictionary<string, string>();
            projekt_path = swModel.GetPathName().Remove(swModel.GetPathName().LastIndexOf((char)92) + 1);

            key = swModel.GetPathName().Substring(swModel.GetPathName().LastIndexOf((char)92) + 1);
            key = key.Substring(0, key.Length - 7);
            pathName = swModel.GetPathName();
            pathName = pathName.Remove(pathName.Length - 7);
            Dict.Add(key, pathName);

            //Создаем список путей компонентов для всех конфигураций
            сonfNames = (string[])swModel.GetConfigurationNames();
            swAssy.ResolveAllLightWeightComponents(false);
            for (int i = 0; i < сonfNames.Length; i++)
            {
                swModel.ShowConfiguration2((string)сonfNames[i]);
                swModel.ForceRebuild3(false);
                Comps = (Object[])swAssy.GetComponents(false);
                for (int j = 0; j < Comps.Length; j++)
                {
                    swComp = (Component2)Comps[j];
                    //compDoc = (ModelDoc2)swComp.GetModelDoc2();
                    if ((swComp.GetSuppression() != (int)swComponentSuppressionState_e.swComponentSuppressed) & (swComp != null))
                    {
                        pathName = swComp.GetPathName();
                        pathName = pathName.Remove(pathName.Length - 7);
                        key = swComp.GetPathName().Substring(swComp.GetPathName().LastIndexOf((char)92) + 1);
                        key = key.Substring(0, key.Length - 7);
                        if (!Dict.ContainsKey(key)) { Dict.Add(key, pathName); }
                    }
                }
            }
            
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

            //Настройки TIF
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffScreenOrPrintCapture, 1); //1-Print capture
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffPrintPaperSize, 12); //12-Papers User Defined
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffImageType, 0); //0-Black And White
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffCompressionScheme, 2); //2-Group 4 Fax Compression
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffPrintDPI, 600); //300-Integer value

            //Сохраняем картинки
            int itogo = 0;

            foreach (KeyValuePair<string, string> k in Drw)
            {
                //Настройка размеров картинки
                iSwApp.IGetTemplateSizes(k.Value + ".SLDDRW", out int PaperSize, out double Width, out double Height);
                iSwApp.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperWidth, Width); //Double value in meters
                iSwApp.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperHeight, Height); //Double value in meters

                Part = (DrawingDoc)iSwApp.OpenDoc6(k.Value + ".SLDDRW", (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_ViewOnly, "", ref errors, ref warnings);
                if ((errors == 0) & (Part != null))
                {
                    swModel = (ModelDoc2)Part;
                    swModelDocExt = (ModelDocExtension)swModel.Extension;
                    swModelDocExt.SaveAs2(projekt_path + "TIF\\" + k.Key + ".TIF", (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent,  null, "", false, ref errors, ref warnings);
                    itogo += 1;
                }
                iSwApp.CloseDoc(k.Value + ".SLDDRW");
                Part = null;
            }
            iSwApp.LoadAddIn(sAddinName);
            iSwApp.SendMsgToUser2("Всего частей " + Dict.Count + System.Environment.NewLine + "Чертежей сохранено " + itogo, 2, 2);
            iSwApp.CommandInProgress = false;
        }
        #endregion

    }


}

