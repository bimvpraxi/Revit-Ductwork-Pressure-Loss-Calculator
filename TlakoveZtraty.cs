// VysokÈ uËenÌ technickÈ v BrnÏ
// ⁄stav automatizace inûen˝rsk˝ch ˙loh a informatiky
// Ing. Michal Nov·Ëek
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
            //Oöet¯enÌ, ûe nelze vybrat jin˝ elemnt neû PÿÕM… POTRUBÕ a POTRUBNÕ TVAROVKA
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
            //CESTA DO PROGRAM FILES PRO TEXç¡K
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

                    // PÿÕPAD, KDY JE V TABULCE VÕCE ÿ¡DKŸ DAT
                    if (File.ReadLines(path).Skip(i + 2).Take(5).First() != "" && File.ReadLines(path).Skip(i).Take(5).First() != "")
                    {
                        linePositionInput = File.ReadLines(path).Skip(i).Take(5).First();

                        // PÿÕPAD KDY R/D JE PÿESNÃ V TABULCE
                        if (linePositionInput.Contains(RD))
                        {
                            charPositionInput = linePositionInput.IndexOf(RD);

                            for (int d = 1; d < maxPocetRadkuVJedneTabulce; d++)
                            {
                                linePositionInputHW = File.ReadLines(path).Skip(i+d).Take(5).First();

                                // PÿÕPAD KDY R/D JE PÿESNÃ V TABULCE A HW SE INTERPOLUJE
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
                                // PÿÕPAD KDY R/D JE PÿESNÃ V TABULCE A HW TAKY
                                if (decimal.Parse(HW.Replace(".", ",")) == decimal.Parse(linePositionInputHW.Substring(0, pocetCislic)))
                                {
                                    coefOutput = linePositionInputHW.Substring(charPositionInput, pocetCislic);
                                    return coefOutput;
                                }
                            }


                        }

                        // PÿÕPAD, KDY R/D JE MEZILEHL¡ HODNOTA A INTERPOLUJE SE
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

                                        // PÿÕPAD KDY HW JE PÿESNÃ V TABULCE A R/D JE INTERPOLOV¡NO
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

                                        // PÿÕPAD, KDY HW NENÕ PÿESNÃ V TABULCE A R/D JE INTERPOLOV¡NO
                                        
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
                           // TaskDialog.Show("Chyba vstupnÌho parametru", "Pro dan˝ pomÏr R/D nebyla nalezena odpovÌdajÌcÌ hodnota koeficientu. R/D je mimo tabel·rnÌ hodnoty");
                        }
                    }

                    // PÿÕPAD KDY JE V TABULCE JENOM JEDEN ÿ¡DEK DAT
                    else
                    {
                        linePositionInput = File.ReadLines(path).Skip(i).Take(5).First();

                        // PÿÕPAD KDY R/D JE PÿESNÃ V TABULCE
                        if (linePositionInput.Contains(RD))
                        {
                            charPositionInput = linePositionInput.IndexOf(RD);
                            linePositionOutput = File.ReadLines(path).Skip(i + 1).Take(5).First();
                            coefOutput = linePositionOutput.Substring(charPositionInput, pocetCislic);
                            return coefOutput;
                        }

                        // PÿÕPAD, KDY R/D JE MEZILEHL¡ HODNOTA A INTERPOLUJE SE
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
            TaskDialog.Show("Chyba souboru", "Chyba v textovÈm souboru");
            return null;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;

            try
            {
                // NastavenÌ vybÌr·nÌ objekt˘. M˘ûou se vybÌrat jenom VZT potrubÌ a pouze Ëtvercov˝m v˝bÏrem
                IList<Element> pickedRef = null;
                Document doc = uiApp.ActiveUIDocument.Document;
                Selection sel = uiApp.ActiveUIDocument.Selection;
                VZTFilter selFilter = new VZTFilter();
                pickedRef = sel.PickElementsByRectangle(selFilter, "Vyberte potrubnÌ systÈm").Select(a => doc.GetElement(a.Id)).Cast<Element>().ToList();

                // Pro kaûd˝ vybran˝ element se zjistÌ hodnota jeho parametru, kter· se uloûÌ do Listu pressureDuct a rozdÏlÌ se na ËÌselnou Ë·st (bez jednotky Pa) a zmÏnÌ se desetinn· teËka za Ë·rku, aby mohl systÈm sËÌtat tyto ËÌsla
                // z LookUp bylo zjiötÏno, ûe parametr m· hodnotu String (AsValueString), proto se hodnoty ukl·dajÌ do stringovÈho Listu
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
                                // ZÌsk·nÌ hodnoty PartType z instance
                                MechanicalFitting mfit = fi.MEPModel as MechanicalFitting;
                                PartType pt = mfit.PartType;
                                string tvarovka = Enum.GetName(typeof(PartType), pt);



                                // KOLENA
                                if (tvarovka == "Elbow")
                                {
                                    radius = fi.LookupParameter("St¯ednÌ polomÏr").AsValueString().Replace(".", ",").Split(' ')[0];

                                    // PODMÕNKA JESTLI SE JEDN¡ O HRANAT… KOLENO NEBO KRUHOV…
                                    //JEDN¡ SE O KRUHOV… KOLENO
                                    if (elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Contains("\u00F8") == true)
                                    {
                                        profile = "round";
                                        HW = "";
                                        angleStr = fi.LookupParameter("⁄hel").AsValueString().Replace(".", ",").Split('∞')[0];

                                        string angleLenghth = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('\u00F8')[0];
                                        
                                        diameter = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('\u00F8')[0];

                                        // ZÌsk·nÌ VelocityPressure z Connector instance
                                        lcon = (fi.MEPModel as MechanicalFitting).ConnectorManager.Connectors.Cast<Connector>().ToList();
                                        c0 = lcon[0];
                                        c1 = lcon[1];

                                        velocityPressure0 = c0.VelocityPressure;
                                        velocityPressure1 = c1.VelocityPressure;

                                        // V˝poËet R/D a zÌsk·nÌ hodnot z textovÈho souboru
                                        decimal RDdbl = decimal.Round(decimal.Parse(radius) / decimal.Parse(diameter), 2, MidpointRounding.AwayFromZero);

                                        string RDstr = RDdbl.ToString();

                                        try
                                        {
                                            // HODNOTY Z TEXç¡KU                                            
                                            double num = double.Parse(GetLine(Enum.GetName(typeof(PartType), pt), profile, RDstr, HW, "Coefficient=0")) * double.Parse(GetLine(Enum.GetName(typeof(PartType), pt), profile, angleStr, HW, "Coefficient=1"));
                                            double pressureLosss = num * velocityPressure0;
                                            pressureDuctFitting.Add(pressureLosss);
                                        }
                                        catch
                                        {
                                        //    TaskDialog.Show("NÏco je öpatnÏ", "Textov˝ soubor nebyl nalezen, ve Visual Studiu prosÌm vyplÚte na ¯·dku 21 spr·vnou cestu k souboru a pomocÌ tlaËÌtka 'Sestavit' a 'Sestavit ¯eöenÌ' vytvo¯Ìte nov˝ *.dll soubor, kter˝ voûÌte do C:\\Users\\XXXXX\\AppData\\Roaming\\Autodesk\\Revit\\Addins\\201Y");
                                        }

                                        // HODNOTY ZE VZTAHU
                                        double num2 = double.Parse(ElbowCalculatePressure(RDstr));
                                        double pressureLosss2 = num2 * velocityPressure0;
                                        pressureDuctFittingEquation.Add(pressureLosss2);
                                    }

                                    // JEDN¡ SE O HRANAT… KOLENO
                                    else
                                    {
                                        profile = "rectangular";
                                        angleStr = fi.LookupParameter("⁄hel").AsValueString().Replace(".", ",").Split('∞')[0];

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

                                        // V›PO»ET KOEFICIENTU TLAKOV›CH ZTR¡T DLE SHWARZERA
                                        string radiusStr = fi.LookupParameter("St¯ednÌ polomÏr").AsValueString().Replace(".", ",").Split(' ')[0];
                                        double radiusDbl = double.Parse(radiusStr);
                                        double ab = Math.Pow(width / height, 2);
                                        double c3 = 1.247 + 0.177 * Math.Log(width / height);
                                        double c2 = 0.092 * Math.Pow(width / height, -1) + 0.046;
                                        double c = (width / height) * Math.Pow(-0.069 - 3.458 * ab, -1);
                                        double rb1 = Math.Pow(radiusDbl / height, -1);
                                        double koef = c * rb1 + c2 * Math.Exp(c3 * rb1);

                                        // V˝poËet R/D a zÌsk·nÌ hodnot z textovÈho souboru
                                        decimal RBdbl = decimal.Round(Convert.ToDecimal(radiusDbl) / Convert.ToDecimal(width), 2, MidpointRounding.AwayFromZero);
                                        string RBstr = RBdbl.ToString();

                                        try
                                        {
                                            // HODNOTY Z TEXç¡KU
                                            double num = double.Parse(GetLine(Enum.GetName(typeof(PartType), pt), profile, RBstr, HW, "Coefficient=0")) * double.Parse(GetLine(Enum.GetName(typeof(PartType), pt), profile, angleStr, "", "Coefficient=1"));
                                            double pressureLosss = num * velocityPressure0;
                                            pressureDuctFitting.Add(pressureLosss);
                                        }
                                        catch
                                        {
                                            TaskDialog.Show("NÏco je öpatnÏ", "Chyba v textovÈm souboru");
                                        }

                                        // HODNOTY ZE VZTAHU DOSAZEN… VYN¡SOBEN… TLAKOVOU RYCHLOSTÕ
                                        double num2 = double.Parse(ElbowCalculatePressure(RBstr));
                                        double pressureLosss2 = num2 * velocityPressure0;
                                        pressureDuctFittingEquation.Add(pressureLosss2);
                                    }
                                }



                                // PÿECHODKY
                                else if (tvarovka == "Transition")
                                {
                                    // KRUHOV¡ PÿECHODKA
                                    if (elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Contains("\u00F8") == true)
                                    {
                                        profile = "round";

                                        diameter = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('\u00F8')[0];
                                        string pomoc_diameter = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('-')[1];
                                        diameter2 = pomoc_diameter.Split('\u00F8')[0];

                                        // V›PO»ET ⁄HLU PÿECHODU
                                        string lengthStr = fi.LookupParameter("VypoËÌtan· dÈlka").AsValueString().Replace(".", ",").Split(' ')[0];
                                        double lengthDbl = double.Parse(lengthStr);
                                        double angleDbl = Math.Atan((Math.Abs(double.Parse(diameter) - double.Parse(diameter2))/2) / lengthDbl) * 180 / Math.PI;
                                        angle = Convert.ToString(angleDbl);

                                        // V›PO»ET PLOCH OBOU STRAN PÿECHODKY
                                        double plocha1 = Math.PI * Math.Pow((double.Parse(diameter) / 2),2);
                                        double plocha2 = Math.PI * Math.Pow((double.Parse(diameter2) / 2), 2);
                                        double podilPlocha = plocha1 / plocha2;

                                        lcon = (fi.MEPModel as MechanicalFitting).ConnectorManager.Connectors.Cast<Connector>().ToList();
                                        c0 = lcon[0];
                                        c1 = lcon[1];

                                        velocityPressure0 = c0.VelocityPressure;
                                        velocityPressure1 = c1.VelocityPressure;

                                        // V›PO»ET POMOCÕ TEXç¡KU
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

                                    // HRNAT¡ PÿECHODKA
                                    else
                                    {
                                        profile = "rectangular";

                                        string pomoc_width = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('-')[0];
                                        widthStr = pomoc_width.Split('x')[0];
                                        heightStr = pomoc_width.Split('x')[1];
                                        string pomoc_width2 = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split('-')[1];
                                        widthStr2 = pomoc_width2.Split('x')[0];
                                        heightStr2 = pomoc_width2.Split('x')[1];

                                        // V›PO»ET ⁄HLU PÿECHODU     
                                            // ZJIäTÃNÕ ZDA SE MÃNÕ V›äKA NEBO äÕÿKA U PÿECHODKY
                                        if (Math.Abs(double.Parse(heightStr) - double.Parse(heightStr2)) == 0)
                                        {
                                            // PÿECHODCE SE MÃNÕ V›äKA
                                            string lengthStr = fi.LookupParameter("VypoËÌtan· dÈlka").AsValueString().Replace(".", ",").Split(' ')[0];
                                            double lengthDbl = double.Parse(lengthStr);
                                            double aa = Math.Abs(double.Parse(widthStr) - double.Parse(widthStr2)) / 2;
                                            double angleDbl = Math.Atan(aa / lengthDbl) * 180 / Math.PI;
                                            decimal angleDec = Convert.ToDecimal(angleDbl);
                                            angleDec = Math.Round(angleDec, 2);
                                            angle = Convert.ToString(angleDec);
                                        }
                                        else
                                        {
                                            // PÿECHODCE SE MÃNÕ äÕÿKA
                                            string lengthStr2 = fi.LookupParameter("VypoËÌtan· dÈlka").AsValueString().Replace(".", ",").Split(' ')[0];
                                            double lengthDbl2 = double.Parse(lengthStr2);
                                            double angleDbl2 = Math.Atan((Math.Abs(double.Parse(heightStr) - double.Parse(heightStr2)) / 2) / lengthDbl2) * 180 / Math.PI;
                                            decimal angleDec2 = Convert.ToDecimal(angleDbl2);
                                            angleDec2 = Math.Round(angleDec2, 2);
                                            angle = Convert.ToString(angleDec2);
                                        }
                                                                               
                                        // V›PO»ET PLOCH OBOU STRAN PÿECHODKY
                                        double plocha1 = double.Parse(widthStr) * double.Parse(heightStr);
                                        double plocha2 = double.Parse(widthStr2) * double.Parse(heightStr2);
                                        double podilPlocha = plocha1 / plocha2;

                                        lcon = (fi.MEPModel as MechanicalFitting).ConnectorManager.Connectors.Cast<Connector>().ToList();
                                        c0 = lcon[0];
                                        c1 = lcon[1];
                                        velocityPressure0 = c0.VelocityPressure;
                                        velocityPressure1 = c1.VelocityPressure;

                                        // V›PO»ET POMOCÕ TEXç¡KU
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
                        TaskDialog.Show("NÏco je öpatnÏ", "Nebyly vybr·ny elementy odpovÌdajÌcÌ vzduchotechnick˝cm prvk˘m s parametry RBS_PRESSURE_DROP a PartType, nebo bylo vybr·no ËtvercovÈ potrubÌ");
                    }
                }

                // ZÌskan˝ parametr nem· vhodn˝ tvar, je pot¯eba z nÏj oddÏlit jednotku a p¯evÈst na double pro dalöÌ v˝poËty + kontrola zdali nenÌ List pr·zdn˝
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

                // SeËtenÌ hodnot v ËÌslenÈm Listu a vyps·nÌ do textovÈho ¯etÏzce         
                StringBuilder st = new StringBuilder();
                string strlossFitting;
                string strlossFitting2;
                string strLossDuct = Math.Round(pressureDuctDbl.Sum(), 3).ToString();

                // Oöet¯enÌ proti pr·zdnÈmu Listu
                if (pressureDuctFitting.Count != 0)
                {
                    st.AppendLine("Hodnoty z textovÈho souboru");
                    st.AppendLine("Tlakov· ztr·ta p¯Ìm˝ch ˙sek˘: " + strLossDuct + " Pa");
                    strlossFitting = Math.Round(pressureDuctFitting.Sum(), 3).ToString();
                    st.AppendLine("Tlakov· ztr·ta tvarovek: " + strlossFitting + " Pa");
                    double dbllossCelk = double.Parse(strLossDuct) + double.Parse(strlossFitting);
                    st.AppendLine("Celkov· tlakov· ztr·ta: " + dbllossCelk.ToString() + " Pa");

                    if (pressureDuctFittingEquation.Count != 0)
                    {
                        st.AppendLine();
                        st.AppendLine("Hodnoty z rovnice");
                        st.AppendLine("Tlakov· ztr·ta p¯Ìm˝ch ˙sek˘: " + strLossDuct + " Pa");
                        strlossFitting2 = Math.Round(pressureDuctFittingEquation.Sum(), 3).ToString();
                        st.AppendLine("Tlakov· ztr·ta tvarovek: " + strlossFitting2 + " Pa");
                        double dbllossCelk2 = double.Parse(strLossDuct) + double.Parse(strlossFitting2);
                        st.AppendLine("Celkov· tlakov· ztr·ta: " + dbllossCelk2.ToString() + " Pa");
                    }
                    else
                    {
                        pressureDuctFittingEquation.Add(0);
                        st.AppendLine();
                        st.AppendLine("Hodnoty z rovnice");
                        st.AppendLine("Tlakov· ztr·ta p¯Ìm˝ch ˙sek˘: " + strLossDuct + " Pa");
                        strlossFitting2 = Math.Round(pressureDuctFittingEquation.Sum(), 3).ToString();
                        st.AppendLine("Tlakov· ztr·ta tvarovek: " + strlossFitting2 + " Pa");
                        double dbllossCelk2 = double.Parse(strLossDuct) + double.Parse(strlossFitting2);
                        st.AppendLine("Celkov· tlakov· ztr·ta: " + dbllossCelk2.ToString() + " Pa");
                    }
                }
                else
                {
                    pressureDuctFitting.Add(0);
                    st.AppendLine("Hodnoty z textovÈho souboru");
                    st.AppendLine("Tlakov· ztr·ta p¯Ìm˝ch ˙sek˘: " + strLossDuct + " Pa");
                    strlossFitting = Math.Round(pressureDuctFitting.Sum(), 3).ToString();
                    st.AppendLine("Tlakov· ztr·ta tvarovek: " + strlossFitting + " Pa");
                    double dbllossCelk = double.Parse(strLossDuct) + double.Parse(strlossFitting);
                    st.AppendLine("Celkov· tlakov· ztr·ta: " + dbllossCelk.ToString() + " Pa");

                    if (pressureDuctFittingEquation.Count != 0)
                    {
                        st.AppendLine();
                        st.AppendLine("Hodnoty z rovnice");
                        st.AppendLine("Tlakov· ztr·ta p¯Ìm˝ch ˙sek˘: " + strLossDuct + " Pa");
                        strlossFitting2 = Math.Round(pressureDuctFittingEquation.Sum(), 3).ToString();
                        double dbllossCelk2 = double.Parse(strLossDuct) + double.Parse(strlossFitting2);
                        st.AppendLine("Celkov· tlakov· ztr·ta: " + dbllossCelk2.ToString() + " Pa");
                    }
                    else
                    {
                        pressureDuctFittingEquation.Add(0);
                        st.AppendLine();
                        st.AppendLine("Hodnoty z rovnice");
                        st.AppendLine("Tlakov· ztr·ta p¯Ìm˝ch ˙sek˘: " + strLossDuct + " Pa");
                        strlossFitting2 = Math.Round(pressureDuctFittingEquation.Sum(), 3).ToString();
                        double dbllossCelk2 = double.Parse(strLossDuct) + double.Parse(strlossFitting2);
                        st.AppendLine("Celkov· tlakov· ztr·ta: " + dbllossCelk2.ToString() + " Pa");
                    }
                }

                // ZobrazenÌ okna s textov˝m ¯etÏzcem v˝sledku
                // PodmÌnka, aby byl vybr·n alespon jeden prvek, jinak systÈm jede d·l (ukonËÌ p¯Ìkaz)
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

                // Jedn·nÌ v p¯ÌpadÏ, ûe p¯Ìkaz bude ukonËen
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

            // Jedn·nÌ v p¯ÌpadÏ, ûe nastane chyba
            catch (Exception ex)
            {
                return Result.Failed;
            }
        }
    }
}
