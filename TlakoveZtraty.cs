// Vysok� u�en� technick� v Brn�
// �stav automatizace in�en�rsk�ch �loh a informatiky
// Ing. Michal Nov��ek
// michal.novace93@seznam.cz
// +420 723 602 636

#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using System.Globalization;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Mechanical;
using System.Security.Policy;
using System.Windows;
#endregion

namespace GGmenu
{
    [Transaction(TransactionMode.Manual)]
    public class TlakoveZtraty : IExternalCommand
    {
        public class VZTFilter : ISelectionFilter
        {
            //O�et�en�, �e nelze vybrat jin� elemnt ne� P��M� POTRUB� a POTRUBN� TVAROVKA
            public bool AllowElement(Element elem)
            {
                if (elem.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_DuctCurves) || elem.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_DuctFitting))
                {
                    return true;
                }
                else return false;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        public string ElbowCalculatePressure(string RD)
        {
            double RDdbl = double.Parse(RD);
            double helpcoef = (9.69 * RDdbl);
            double helpcoef2 = helpcoef - 4.24;

            double helpcoef3 = Math.Pow(helpcoef2, -1);
            double Eta = RDdbl * helpcoef3;
            return Eta.ToString();
        }

        public string TransitionCalculatePressure(string areas)
        {
            double RDdbl = double.Parse(areas);
            double helpcoef = double.Parse(areas);
            double Eta = Math.Pow((helpcoef - 1),2);
            return Eta.ToString();
        }

        public string GetLine(string partType, string profile, string RD, string HW, string coef)
        {
            //CESTA DO PROGRAM FILES PRO TEX��K
            var domaininfo = new AppDomainSetup();
            Evidence adEvidence = AppDomain.CurrentDomain.Evidence;
            AppDomain domain = AppDomain.CreateDomain("Domain2", adEvidence, domaininfo);
            string path = domain.BaseDirectory + "\\Koeficienty.txt";

            string coefOutput;
            string coefOutputI1;
            string line;
            string linePositionInput;
            string linePositionInputHW;
            string linePositionOutput;
            string linePositionOutput2;
            int charPositionInput;
            int pocetCislic = 6;
            int pocetZnakuMeziZarazkami = 8;
            int maxPocetRadkuVJedneTabulce = 100;
            int maxPocetZnakuNaRadku = 150;

            if (RD.Length != pocetCislic)
            {
                if (RD.Contains(","))
                {
                    for (int i = 0; i < (pocetCislic - RD.Length + 1); i++)
                    {
                        RD = RD + "0";
                    }
                }
                else
                {
                    RD = RD + ",";
                    for (int i = 0; i < (pocetCislic - RD.Length+2); i++)
                    {
                        RD = RD + "0";
                    }
                }
            }

            RD = " " + RD;

            //    RD = " " + RD + " ";

            // Read the file and display it line by line.
            StreamReader file = new StreamReader(path);
            for (int i = 1; i<File.ReadLines(path).Count(); i++)
            {
                line = file.ReadLine();
                if (line.Contains(partType) && line.Contains(profile) && line.Contains(coef))
                {

                    // P��PAD, KDY JE V TABULCE V�CE ��DK� DAT
                    if (File.ReadLines(path).Skip(i + 2).Take(5).First() != "" && File.ReadLines(path).Skip(i).Take(5).First() != "")
                    {
                        linePositionInput = File.ReadLines(path).Skip(i).Take(5).First();

                        // P��PAD KDY R/D JE P�ESN� V TABULCE
                        if (linePositionInput.Contains(RD))
                        {
                            charPositionInput = linePositionInput.IndexOf(RD);

                            for (int d = 1; d < maxPocetRadkuVJedneTabulce; d++)
                            {
                                linePositionInputHW = File.ReadLines(path).Skip(i+d).Take(5).First();

                                // P��PAD KDY R/D JE P�ESN� V TABULCE A HW SE INTERPOLUJE
                                if (decimal.Parse(HW.Replace(".", ",")) < decimal.Parse(linePositionInputHW.Substring(0, pocetCislic)) && decimal.Parse(HW.Replace(".", ",")) != decimal.Parse(linePositionInputHW.Substring(0, pocetCislic)))
                                {
                                    linePositionInput = File.ReadLines(path).Skip(i + d - 1).Take(5).First();
                                    string xstr = linePositionInput.Substring(0, pocetCislic);
                                    string xstr2 = linePositionInputHW.Substring(0, pocetCislic);

                                    string ystr = linePositionInput.Substring(charPositionInput, pocetCislic);
                                    string ystr2 = linePositionInputHW.Substring(charPositionInput, pocetCislic);

                                    

                                    double x = double.Parse(xstr);
                                    double x2 = double.Parse(xstr2);
                                    double y = double.Parse(ystr);
                                    double y2 = double.Parse(ystr2);
                                    double vys = (y2 - y) * (double.Parse(HW) - x) / (x2 - x) + y;

                                    coefOutput = Convert.ToString(vys);
                                    return coefOutput;
                                }
                                // P��PAD KDY R/D JE P�ESN� V TABULCE A HW TAKY
                                if (decimal.Parse(HW.Replace(".", ",")) == decimal.Parse(linePositionInputHW.Substring(0, pocetCislic)))
                                {
                                    coefOutput = linePositionInputHW.Substring(charPositionInput, pocetCislic);
                                    return coefOutput;
                                }
                            }


                        }

                        // P��PAD, KDY R/D JE MEZILEHL� HODNOTA A INTERPOLUJE SE
                        else
                        {
                            linePositionOutput = File.ReadLines(path).Skip(i).Take(5).First();
                            linePositionOutput2 = File.ReadLines(path).Skip(i-1).Take(5).First();
                            for (int h = pocetZnakuMeziZarazkami-1; h < maxPocetZnakuNaRadku; h = h + pocetZnakuMeziZarazkami)
                            {
                                if (decimal.Parse(RD.Replace(".", ",")) < decimal.Parse(linePositionOutput.Substring(h, pocetCislic)))
                                {
                                    for (int d = 1; i < maxPocetRadkuVJedneTabulce; i++)
                                    {
                                        linePositionInputHW = File.ReadLines(path).Skip(i+d).Take(5).First();

                                        // P��PAD KDY HW JE P�ESN� V TABULCE A R/D JE INTERPOLOV�NO
                                        if (decimal.Parse(HW.Replace(".", ",")) == decimal.Parse(linePositionInputHW.Substring(0, pocetCislic)))
                                        {
                                            linePositionInputHW = File.ReadLines(path).Skip(i + d).Take(5).First();
                                            string x2str;
                                            double x2;
                                            string y2str;
                                            double y2;
                                            string x1str;
                                            double x1;
                                            string y1str;
                                            double y1;
                                            double x;
                                            double y;
                                            string nextLine;

                                            x = double.Parse(RD);
                                            x1str = linePositionOutput.Substring(h - pocetZnakuMeziZarazkami, pocetCislic);
                                            x2str = linePositionOutput.Substring(h, pocetCislic);
                                            nextLine = File.ReadLines(path).Skip(i + d).Take(5).First();
                                            y1str = linePositionInputHW.Substring(h - pocetZnakuMeziZarazkami, pocetCislic);
                                            y2str = linePositionInputHW.Substring(h, pocetCislic);

                                            x1 = double.Parse(x1str);
                                            y1 = double.Parse(y1str);
                                            x2 = double.Parse(x2str);
                                            y2 = double.Parse(y2str);
                                            y = (y2 - y1) * (x - x1) / (x2 - x1) + y1;
                                            coefOutputI1 = Convert.ToString(y);
                                            return coefOutputI1;
                                        }

                                        // P��PAD, KDY HW NEN� P�ESN� V TABULCE A R/D JE INTERPOLOV�NO
                                        
                                        if (decimal.Parse(HW.Replace(".", ",")) < decimal.Parse(linePositionInputHW.Substring(0, pocetCislic)) && decimal.Parse(HW.Replace(".", ",")) != decimal.Parse(linePositionInputHW.Substring(0, pocetCislic)))
                                        {
                                            linePositionInputHW = File.ReadLines(path).Skip(i + d - 1).Take(5).First();
                                            string x2str;
                                            double x2;
                                            string y2str;
                                            string y2str2;
                                            double y2;

                                            double Vysy2;
                                            double Vysy;

                                            double y2_2;
                                            string x1str;
                                            double x1;
                                            string y1str;
                                            string y1str2;
                                            double y1;
                                            double y1_2;

                                            double x;

                                            string nextLine;
                                            x = double.Parse(RD);
                                            x1str = linePositionOutput.Substring(h - pocetZnakuMeziZarazkami, pocetCislic);
                                            x2str = linePositionOutput.Substring(h, pocetCislic);

                                            nextLine = File.ReadLines(path).Skip(i + d).Take(5).First();
                                            y1str = linePositionInputHW.Substring(h - pocetZnakuMeziZarazkami, pocetCislic);
                                            y2str = linePositionInputHW.Substring(h, pocetCislic);

                                            y1str2 = nextLine.Substring(h - pocetZnakuMeziZarazkami, pocetCislic);
                                            y2str2 = nextLine.Substring(h, pocetCislic);

                                            x1 = double.Parse(x1str);
                                            y1 = double.Parse(y1str);
                                            y1_2 = double.Parse(y1str2);
                                            x2 = double.Parse(x2str);
                                            y2 = double.Parse(y2str);
                                            y2_2 = double.Parse(y2str2);

                                            Vysy = (y2 - y1) * (x - x1) / (x2 - x1) + y1;
                                            Vysy2 = (y2_2 - y1_2) * (x - x1) / (x2 - x1) + y1_2;

                                            double vys;

                                            string HWy1;
                                            string HWy2;
                                            double HWy1Dbl;
                                            double HWy2Dbl;

                                            HWy1 = linePositionInputHW = File.ReadLines(path).Skip(i + d - 1).Take(5).First().Substring(0, pocetCislic);
                                            HWy2 = linePositionInputHW = File.ReadLines(path).Skip(i + d).Take(5).First().Substring(0, pocetCislic);

                                            HWy1Dbl = double.Parse(HWy1);
                                            HWy2Dbl = double.Parse(HWy2);

                                            vys = (Vysy2 - Vysy) * (double.Parse(HW) - HWy1Dbl) / (HWy2Dbl - HWy1Dbl) + Vysy;

                                            coefOutputI1 = Convert.ToString(vys);
                                            return coefOutputI1;
                                        }
                                    }
                                }
                            }
                           // TaskDialog.Show("Chyba vstupn�ho parametru", "Pro dan� pom�r R/D nebyla nalezena odpov�daj�c� hodnota koeficientu. R/D je mimo tabel�rn� hodnoty");
                        }
                    }

                    // P��PAD KDY JE V TABULCE JENOM JEDEN ��DEK DAT
                    else
                    {
                        linePositionInput = File.ReadLines(path).Skip(i).Take(5).First();

                        // P��PAD KDY R/D JE P�ESN� V TABULCE
                        if (linePositionInput.Contains(RD))
                        {
                            charPositionInput = linePositionInput.IndexOf(RD);
                            linePositionOutput = File.ReadLines(path).Skip(i + 1).Take(5).First();
                            coefOutput = linePositionOutput.Substring(charPositionInput, pocetCislic);
                            return coefOutput;
                        }

                        // P��PAD, KDY R/D JE MEZILEHL� HODNOTA A INTERPOLUJE SE
                        else
                        {
                            linePositionOutput = File.ReadLines(path).Skip(i).Take(5).First();
                            for (int h = pocetZnakuMeziZarazkami-1; h < maxPocetZnakuNaRadku; h = h + pocetZnakuMeziZarazkami)
                            {
                                if (decimal.Parse(RD.Replace(".", ",")) < decimal.Parse(linePositionOutput.Substring(h, pocetCislic)))
                                {
                                    string x2str;
                                    double x2;
                                    string y2str;
                                    double y2;
                                    string x1str;
                                    double x1;
                                    string y1str;
                                    double y1;

                                    double x;

                                    double y;
                                    string nextLine;
                                    x = double.Parse(RD);
                                    x1str = linePositionOutput.Substring(h - pocetZnakuMeziZarazkami, pocetCislic);
                                    x2str = linePositionOutput.Substring(h, pocetCislic);
                                    nextLine = File.ReadLines(path).Skip(i + 1).Take(5).First();
                                    y1str = nextLine.Substring(h - pocetZnakuMeziZarazkami, pocetCislic);
                                    y2str = nextLine.Substring(h, pocetCislic);

                                    x1 = double.Parse(x1str);
                                    y1 = double.Parse(y1str);
                                    x2 = double.Parse(x2str);
                                    y2 = double.Parse(y2str);

                                    y = (y2 - y1) * (x - x1) / (x2 - x1) + y1;
                                    coefOutput = Convert.ToString(y);
                                    return coefOutput;
                                }
                            }
                        }                
                    }
                }
                else continue;
            }
            TaskDialog.Show("Chyba souboru", "Chyba v textov�m souboru");
            return null;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;

            try
            {
                // Nastaven� vyb�r�n� objekt�. M��ou se vyb�rat jenom VZT potrub� a pouze �tvercov�m v�b�rem
                IList<Element> pickedRef = null;
                Document doc = uiApp.ActiveUIDocument.Document;
                Selection sel = uiApp.ActiveUIDocument.Selection;
                VZTFilter selFilter = new VZTFilter();
                pickedRef = sel.PickElementsByRectangle(selFilter, "Vyberte potrubn� syst�m").Select(a => doc.GetElement(a.Id)).Cast<Element>().ToList();

                // Pro ka�d� vybran� element se zjist� hodnota jeho parametru, kter� se ulo�� do Listu pressureDuct a rozd�l� se na ��selnou ��st (bez jednotky Pa) a zm�n� se desetinn� te�ka za ��rku, aby mohl syst�m s��tat tyto ��sla
                // z LookUp bylo zji�t�no, �e parametr m� hodnotu String (AsValueString), proto se hodnoty ukl�daj� do stringov�ho Listu
                List<string> pressureDuct = new List<string>();
                List<double> pressureDuctFitting = new List<double>();
                List<double> pressureDuctFittingEquation = new List<double>();
                List<Connector> lcon;

                Connector c0;
                Connector c1;

                double velocityPressure0;
                double velocityPressure1;

                string radius;
                string diameter;
                string diameter2;
                string angle;

                string widthStr;
                string heightStr;
                double width;
                double height;
                string angleStr;

                string widthStr2;
                string heightStr2;

                string pomocne;

                string profile;

                string HW;
                double HWdbl;

                foreach (Element elem in pickedRef)
                {
                    try
                    {
                        if (elem is FamilyInstance fi)
                        {
                            if (fi.MEPModel is MechanicalFitting)
                            {
                                // Z�sk�n� hodnoty PartType z instance
                                MechanicalFitting mfit = fi.MEPModel as MechanicalFitting;
                                PartType pt = mfit.PartType;
                                string tvarovka = Enum.GetName(typeof(PartType), pt);



                                // KOLENA
                                if (tvarovka == "Elbow")
                                {
                                    radius = fi.LookupParameter("St�edn� polom�r").AsValueString().Replace(".", ",").Split(' ')[0];

                                    // PODM�NKA JESTLI SE JEDN� O HRANAT� KOLENO NEBO KRUHOV�
                                    //JEDN� SE O KRUHOV� KOLENO
                                    if (elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Contains("\u00F8") == true)
                                    {
                                        profile = "round";
                                        HW = "";
                                        angleStr = fi.LookupParameter("�hel").AsValueString().Replace(".", ",").Split('�')[0];

                                        string angleLenghth = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('\u00F8')[0];
                                        
                                        diameter = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('\u00F8')[0];

                                        // Z�sk�n� VelocityPressure z Connector instance
                                        lcon = (fi.MEPModel as MechanicalFitting).ConnectorManager.Connectors.Cast<Connector>().ToList();
                                        c0 = lcon[0];
                                        c1 = lcon[1];

                                        velocityPressure0 = c0.VelocityPressure;
                                        velocityPressure1 = c1.VelocityPressure;

                                        // V�po�et R/D a z�sk�n� hodnot z textov�ho souboru
                                        decimal RDdbl = decimal.Round(decimal.Parse(radius) / decimal.Parse(diameter), 2, MidpointRounding.AwayFromZero);

                                        string RDstr = RDdbl.ToString();

                                        try
                                        {
                                            // HODNOTY Z TEX��KU                                            
                                            double num = double.Parse(GetLine(Enum.GetName(typeof(PartType), pt), profile, RDstr, HW, "Coefficient=0")) * double.Parse(GetLine(Enum.GetName(typeof(PartType), pt), profile, angleStr, HW, "Coefficient=1"));
                                            double pressureLosss = num * velocityPressure0;
                                            pressureDuctFitting.Add(pressureLosss);
                                        }
                                        catch
                                        {
                                        //    TaskDialog.Show("N�co je �patn�", "Textov� soubor nebyl nalezen, ve Visual Studiu pros�m vypl�te na ��dku 21 spr�vnou cestu k souboru a pomoc� tla��tka 'Sestavit' a 'Sestavit �e�en�' vytvo��te nov� *.dll soubor, kter� vo��te do C:\\Users\\XXXXX\\AppData\\Roaming\\Autodesk\\Revit\\Addins\\201Y");
                                        }

                                        // HODNOTY ZE VZTAHU
                                        double num2 = double.Parse(ElbowCalculatePressure(RDstr));
                                        double pressureLosss2 = num2 * velocityPressure0;
                                        pressureDuctFittingEquation.Add(pressureLosss2);
                                    }

                                    // JEDN� SE O HRANAT� KOLENO
                                    else
                                    {
                                        profile = "rectangular";
                                        angleStr = fi.LookupParameter("�hel").AsValueString().Replace(".", ",").Split('�')[0];

                                        pomocne = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('-')[0];
                                        widthStr = pomocne.Split('x')[0];
                                        heightStr = pomocne.Split('x')[1];

                                        width = double.Parse(widthStr);
                                        height = double.Parse(heightStr);

                                        HWdbl = height / width;
                                        HW = Convert.ToString(HWdbl);

                                        lcon = (fi.MEPModel as MechanicalFitting).ConnectorManager.Connectors.Cast<Connector>().ToList();
                                        c0 = lcon[0];
                                        c1 = lcon[1];

                                        velocityPressure0 = c0.VelocityPressure;
                                        velocityPressure1 = c1.VelocityPressure;

                                        // V�PO�ET KOEFICIENTU TLAKOV�CH ZTR�T DLE SHWARZERA
                                        string radiusStr = fi.LookupParameter("St�edn� polom�r").AsValueString().Replace(".", ",").Split(' ')[0];
                                        double radiusDbl = double.Parse(radiusStr);
                                        double ab = Math.Pow(width / height, 2);
                                        double c3 = 1.247 + 0.177 * Math.Log(width / height);
                                        double c2 = 0.092 * Math.Pow(width / height, -1) + 0.046;
                                        double c = (width / height) * Math.Pow(-0.069 - 3.458 * ab, -1);
                                        double rb1 = Math.Pow(radiusDbl / height, -1);
                                        double koef = c * rb1 + c2 * Math.Exp(c3 * rb1);

                                        // V�po�et R/D a z�sk�n� hodnot z textov�ho souboru
                                        decimal RBdbl = decimal.Round(Convert.ToDecimal(radiusDbl) / Convert.ToDecimal(width), 2, MidpointRounding.AwayFromZero);
                                        string RBstr = RBdbl.ToString();

                                        try
                                        {
                                            // HODNOTY Z TEX��KU
                                            double num = double.Parse(GetLine(Enum.GetName(typeof(PartType), pt), profile, RBstr, HW, "Coefficient=0")) * double.Parse(GetLine(Enum.GetName(typeof(PartType), pt), profile, angleStr, "", "Coefficient=1"));
                                            double pressureLosss = num * velocityPressure0;
                                            pressureDuctFitting.Add(pressureLosss);
                                        }
                                        catch
                                        {
                                            TaskDialog.Show("N�co je �patn�", "Chyba v textov�m souboru");
                                        }

                                        // HODNOTY ZE VZTAHU DOSAZEN� VYN�SOBEN� TLAKOVOU RYCHLOST�
                                        double num2 = double.Parse(ElbowCalculatePressure(RBstr));
                                        double pressureLosss2 = num2 * velocityPressure0;
                                        pressureDuctFittingEquation.Add(pressureLosss2);
                                    }
                                }



                                // P�ECHODKY
                                else if (tvarovka == "Transition")
                                {
                                    // KRUHOV� P�ECHODKA
                                    if (elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Contains("\u00F8") == true)
                                    {
                                        profile = "round";

                                        diameter = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('\u00F8')[0];
                                        string pomoc_diameter = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('-')[1];
                                        diameter2 = pomoc_diameter.Split('\u00F8')[0];

                                        // V�PO�ET �HLU P�ECHODU
                                        string lengthStr = fi.LookupParameter("Vypo��tan� d�lka").AsValueString().Replace(".", ",").Split(' ')[0];
                                        double lengthDbl = double.Parse(lengthStr);
                                        double angleDbl = Math.Atan((Math.Abs(double.Parse(diameter) - double.Parse(diameter2))/2) / lengthDbl) * 180 / Math.PI;
                                        angle = Convert.ToString(angleDbl);

                                        // V�PO�ET PLOCH OBOU STRAN P�ECHODKY
                                        double plocha1 = Math.PI * Math.Pow((double.Parse(diameter) / 2),2);
                                        double plocha2 = Math.PI * Math.Pow((double.Parse(diameter2) / 2), 2);
                                        double podilPlocha = plocha1 / plocha2;

                                        lcon = (fi.MEPModel as MechanicalFitting).ConnectorManager.Connectors.Cast<Connector>().ToList();
                                        c0 = lcon[0];
                                        c1 = lcon[1];

                                        velocityPressure0 = c0.VelocityPressure;
                                        velocityPressure1 = c1.VelocityPressure;

                                        // V�PO�ET POMOC� TEX��KU
                                        double num = double.Parse(GetLine(Enum.GetName(typeof(PartType), pt), profile, angle, Convert.ToString(podilPlocha), "Coefficient=0"));
                                        double pressureLosss = num * velocityPressure0;
                                        pressureDuctFitting.Add(pressureLosss);

                                        // HODNOTY ZE VZTAHU
                                        double area1 = Math.PI * Math.Pow(double.Parse(diameter), 2) / 4;
                                        double area2 = Math.PI * Math.Pow(double.Parse(diameter2), 2) / 4;
                                        double areas = area2 / area1;

                                        double num2 = double.Parse(TransitionCalculatePressure(Convert.ToString(areas)));
                                        double pressureLosss2 = num2 * velocityPressure0;
                                        pressureDuctFittingEquation.Add(pressureLosss2);
                                    }

                                    // HRNAT� P�ECHODKA
                                    else
                                    {
                                        profile = "rectangular";

                                        string pomoc_width = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('-')[0];
                                        widthStr = pomoc_width.Split('x')[0];
                                        heightStr = pomoc_width.Split('x')[1];
                                        string pomoc_width2 = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('-')[1];
                                        widthStr2 = pomoc_width2.Split('x')[0];
                                        heightStr2 = pomoc_width2.Split('x')[1];

                                        // V�PO�ET �HLU P�ECHODU     
                                            // ZJI�T�N� ZDA SE M�N� V݊KA NEBO ���KA U P�ECHODKY
                                        if (Math.Abs(double.Parse(heightStr) - double.Parse(heightStr2)) == 0)
                                        {
                                            // P�ECHODCE SE M�N� V݊KA
                                            string lengthStr = fi.LookupParameter("Vypo��tan� d�lka").AsValueString().Replace(".", ",").Split(' ')[0];
                                            double lengthDbl = double.Parse(lengthStr);
                                            double aa = Math.Abs(double.Parse(widthStr) - double.Parse(widthStr2)) / 2;
                                            double angleDbl = Math.Atan(aa / lengthDbl) * 180 / Math.PI;
                                            decimal angleDec = Convert.ToDecimal(angleDbl);
                                            angleDec = Math.Round(angleDec, 2);
                                            angle = Convert.ToString(angleDec);
                                        }
                                        else
                                        {
                                            // P�ECHODCE SE M�N� ���KA
                                            string lengthStr2 = fi.LookupParameter("Vypo��tan� d�lka").AsValueString().Replace(".", ",").Split(' ')[0];
                                            double lengthDbl2 = double.Parse(lengthStr2);
                                            double angleDbl2 = Math.Atan((Math.Abs(double.Parse(heightStr) - double.Parse(heightStr2)) / 2) / lengthDbl2) * 180 / Math.PI;
                                            decimal angleDec2 = Convert.ToDecimal(angleDbl2);
                                            angleDec2 = Math.Round(angleDec2, 2);
                                            angle = Convert.ToString(angleDec2);
                                        }
                                                                               
                                        // V�PO�ET PLOCH OBOU STRAN P�ECHODKY
                                        double plocha1 = double.Parse(widthStr) * double.Parse(heightStr);
                                        double plocha2 = double.Parse(widthStr2) * double.Parse(heightStr2);
                                        double podilPlocha = plocha1 / plocha2;

                                        lcon = (fi.MEPModel as MechanicalFitting).ConnectorManager.Connectors.Cast<Connector>().ToList();
                                        c0 = lcon[0];
                                        c1 = lcon[1];
                                        velocityPressure0 = c0.VelocityPressure;
                                        velocityPressure1 = c1.VelocityPressure;

                                        // V�PO�ET POMOC� TEX��KU
                                        double num = double.Parse(GetLine(Enum.GetName(typeof(PartType), pt), profile, angle, Convert.ToString(podilPlocha), "Coefficient=0"));
                                        double pressureLosss = num * velocityPressure0;
                                        pressureDuctFitting.Add(pressureLosss);

                                        // HODNOTY ZE VZTAHU
                                        double area1 = double.Parse(widthStr) * double.Parse(heightStr);
                                        double area2 = double.Parse(widthStr2) * double.Parse(heightStr2);
                                        double areas = area2 / area1;
                                        double num2 = double.Parse(TransitionCalculatePressure(Convert.ToString(areas)));
                                        double pressureLosss2 = num2 * velocityPressure0;
                                        pressureDuctFittingEquation.Add(pressureLosss2);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (elem is Duct d)
                            {
                                pressureDuct.Add(elem.get_Parameter(BuiltInParameter.RBS_PRESSURE_DROP).AsValueString().Replace(".", ",").Split(' ')[0]);
                            }
                        }
                    }
                    catch
                    {
                        TaskDialog.Show("N�co je �patn�", "Nebyly vybr�ny elementy odpov�daj�c� vzduchotechnick�cm prvk�m s parametry RBS_PRESSURE_DROP a PartType, nebo bylo vybr�no �tvercov� potrub�");
                    }
                }

                // Z�skan� parametr nem� vhodn� tvar, je pot�eba z n�j odd�lit jednotku a p�ev�st na double pro dal�� v�po�ty + kontrola zdali nen� List pr�zdn�
                List<double> pressureDuctDbl = new List<double>();
                if (pressureDuct.Count != 0)
                {
                    for (int i = 0; i < pressureDuct.Count; i++)
                    {
                        double split = double.Parse(pressureDuct[i]);
                        pressureDuctDbl.Add(split);
                    }
                }
                else
                {
                    pressureDuctDbl.Add(0);
                }

                // Se�ten� hodnot v ��slen�m Listu a vyps�n� do textov�ho �et�zce         
                StringBuilder st = new StringBuilder();
                string strlossFitting;
                string strlossFitting2;
                string strLossDuct = Math.Round(pressureDuctDbl.Sum(), 3).ToString();

                // O�et�en� proti pr�zdn�mu Listu
                if (pressureDuctFitting.Count != 0)
                {
                    st.AppendLine("Hodnoty z textov�ho souboru");
                    st.AppendLine("Tlakov� ztr�ta p��m�ch �sek�: " + strLossDuct + " Pa");
                    strlossFitting = Math.Round(pressureDuctFitting.Sum(), 3).ToString();
                    st.AppendLine("Tlakov� ztr�ta tvarovek: " + strlossFitting + " Pa");
                    double dbllossCelk = double.Parse(strLossDuct) + double.Parse(strlossFitting);
                    st.AppendLine("Celkov� tlakov� ztr�ta: " + dbllossCelk.ToString() + " Pa");

                    if (pressureDuctFittingEquation.Count != 0)
                    {
                        st.AppendLine();
                        st.AppendLine("Hodnoty z rovnice");
                        st.AppendLine("Tlakov� ztr�ta p��m�ch �sek�: " + strLossDuct + " Pa");
                        strlossFitting2 = Math.Round(pressureDuctFittingEquation.Sum(), 3).ToString();
                        st.AppendLine("Tlakov� ztr�ta tvarovek: " + strlossFitting2 + " Pa");
                        double dbllossCelk2 = double.Parse(strLossDuct) + double.Parse(strlossFitting2);
                        st.AppendLine("Celkov� tlakov� ztr�ta: " + dbllossCelk2.ToString() + " Pa");
                    }
                    else
                    {
                        pressureDuctFittingEquation.Add(0);
                        st.AppendLine();
                        st.AppendLine("Hodnoty z rovnice");
                        st.AppendLine("Tlakov� ztr�ta p��m�ch �sek�: " + strLossDuct + " Pa");
                        strlossFitting2 = Math.Round(pressureDuctFittingEquation.Sum(), 3).ToString();
                        st.AppendLine("Tlakov� ztr�ta tvarovek: " + strlossFitting2 + " Pa");
                        double dbllossCelk2 = double.Parse(strLossDuct) + double.Parse(strlossFitting2);
                        st.AppendLine("Celkov� tlakov� ztr�ta: " + dbllossCelk2.ToString() + " Pa");
                    }
                }
                else
                {
                    pressureDuctFitting.Add(0);
                    st.AppendLine("Hodnoty z textov�ho souboru");
                    st.AppendLine("Tlakov� ztr�ta p��m�ch �sek�: " + strLossDuct + " Pa");
                    strlossFitting = Math.Round(pressureDuctFitting.Sum(), 3).ToString();
                    st.AppendLine("Tlakov� ztr�ta tvarovek: " + strlossFitting + " Pa");
                    double dbllossCelk = double.Parse(strLossDuct) + double.Parse(strlossFitting);
                    st.AppendLine("Celkov� tlakov� ztr�ta: " + dbllossCelk.ToString() + " Pa");

                    if (pressureDuctFittingEquation.Count != 0)
                    {
                        st.AppendLine();
                        st.AppendLine("Hodnoty z rovnice");
                        st.AppendLine("Tlakov� ztr�ta p��m�ch �sek�: " + strLossDuct + " Pa");
                        strlossFitting2 = Math.Round(pressureDuctFittingEquation.Sum(), 3).ToString();
                        double dbllossCelk2 = double.Parse(strLossDuct) + double.Parse(strlossFitting2);
                        st.AppendLine("Celkov� tlakov� ztr�ta: " + dbllossCelk2.ToString() + " Pa");
                    }
                    else
                    {
                        pressureDuctFittingEquation.Add(0);
                        st.AppendLine();
                        st.AppendLine("Hodnoty z rovnice");
                        st.AppendLine("Tlakov� ztr�ta p��m�ch �sek�: " + strLossDuct + " Pa");
                        strlossFitting2 = Math.Round(pressureDuctFittingEquation.Sum(), 3).ToString();
                        double dbllossCelk2 = double.Parse(strLossDuct) + double.Parse(strlossFitting2);
                        st.AppendLine("Celkov� tlakov� ztr�ta: " + dbllossCelk2.ToString() + " Pa");
                    }
                }

                // Zobrazen� okna s textov�m �et�zcem v�sledku
                // Podm�nka, aby byl vybr�n alespon jeden prvek, jinak syst�m jede d�l (ukon�� p��kaz)
                    if (pickedRef.Count != 0)
                    {
                        TaskDialog.Show("Pressure loss", st.ToString());
                    }
                    else
                    {
                        return Result.Cancelled;
                    }

                    return Result.Succeeded;
                    }

                // Jedn�n� v p��pad�, �e p��kaz bude ukon�en
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

            // Jedn�n� v p��pad�, �e nastane chyba
            catch (Exception ex)
            {
                return Result.Failed;
            }
        }
    }
}
