using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace PAC_Sheets
{
    class ExcelDataManager
    {
        XLWorkbook workbook = new XLWorkbook(System.AppDomain.CurrentDomain.BaseDirectory + "Template.xlsx");
        int worksheetIndex = 1;
        string siteName;

        public void InsertData(string[,] strategyData)
        {
            try
            {
                siteName = strategyData[0, 5];
                foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                {
                    siteName = siteName.Replace(c, '-');
                }
                int i = 0;
                int rowIndex = 9;
                string description = strategyData[2, i];//Base IO, 8UI, 16DI 1, etc.
                string IORef = "";
                if (description != "") { IORef = description; IORef = strategyData[5, i] + " " + strategyData[18, i]; }
                IXLWorksheet currentSheet = workbook.Worksheet(worksheetIndex);
                string aWName = strategyData[2, i];
                //string aWName = strategyData[5, i] +  " " + strategyData[18, i];
                foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                {
                    aWName = aWName.Replace(c, '-');
                }
                currentSheet.Name = aWName;
                //copy style
                var style1 = currentSheet.Cell(8, 1).Style;
                while (description != null)
                {
                    if (description != "")//new or first sheet
                    {
                        currentSheet = workbook.Worksheet(worksheetIndex);
                        //string tabName = strategyData[2, i];
                        string tabName = "";//Need to make sure it does not contain any ilegal characters :\\/?*[]"
                        if (strategyData[5, i] == "0") { tabName = strategyData[18, i]; } else { tabName = strategyData[5, i] + " " + strategyData[18, i];}
                        //FIX HERE*************************************************************************************************
                        if (tabName == "") { tabName = strategyData[2, i]; }
                        foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                        {
                            tabName = tabName.Replace(c, '-');
                        }
                        currentSheet.Name = tabName;
                        currentSheet.Cell(5, 4).Value = strategyData[0, 0];
                        currentSheet.Cell(5, 2).Value = strategyData[0, 1];
                        currentSheet.Cell(4, 10).Value = strategyData[0, 2];
                        currentSheet.Cell(4, 8).Value = strategyData[0, 3];
                        currentSheet.Cell(4, 6).Value = strategyData[0, 4];
                        currentSheet.Cell(4, 2).Value = strategyData[0, 5];
                        currentSheet.Cell(5, 10).Value = strategyData[0, 6];

                        // Write first point
                        //Channel
                        currentSheet.Cell(8, 1).Value = strategyData[1, i];
                        //Inputs/output
                        if (worksheetIndex == 1)
                        {
                            if (strategyData[3, 0].Contains("I")) { currentSheet.Cell(8, 2).Value = "Inputs"; } else { currentSheet.Cell(8, 2).Value = "Outputs"; }
                        }
                        else
                        {
                            currentSheet.Cell(8, 2).Value = description;
                            currentSheet.Cell(8, 2).Value = strategyData[5, i] + " " + strategyData[18, i];
                        }
                        //Labels and types
                        if (strategyData[4, i].Contains("S"))// get sensor label and type - data[8, i]=sensorNumber, data[9, i]=sensorLabel, data[10, i]=sensorType
                        {
                            string sensorNumber = strategyData[4, i].Substring(1, strategyData[4, i].Length - 1);
                            int s = 0;
                            while (strategyData[8, s] != sensorNumber) { s++; }
                            currentSheet.Cell(8, 4).Value = strategyData[9, s];
                            string sensorType = strategyData[10, s];//data[15, i]=typeNumbers, data[16, i]=setPartNumbers
                            int t = 0;
                            while (strategyData[15, t] != sensorType) { t++; }
                            currentSheet.Cell(8, 3).Value = strategyData[16, t];

                        }
                        if (strategyData[4, i].Contains("D"))// get driver label and type (digital or analogue, etc.) - data[13, i]=driverNumber, data[14, i]=driverLabel
                        {
                            string driverNumber = strategyData[4, i].Substring(1, strategyData[4, i].Length - 2);
                            int d = 0;
                            while (strategyData[13, d] != driverNumber) { d++; }
                            currentSheet.Cell(8, 4).Value = strategyData[14, d];
                            string driverType = strategyData[17, d];
                            if (driverType == "1") { currentSheet.Cell(8, 3).Value = "Digital Driver"; }
                            if (driverType == "2") { currentSheet.Cell(8, 3).Value = "Analogue Driver"; }
                            if (driverType == "3") { currentSheet.Cell(8, 3).Value = "Time Proportional"; }
                            if (driverType == "4") { currentSheet.Cell(8, 3).Value = "Raise/Lower End"; }
                            if (driverType == "5") { currentSheet.Cell(8, 3).Value = "Binary Histeresis"; }
                            if (driverType == "6") { currentSheet.Cell(8, 3).Value = "Time Proportional +O/R"; }
                            if (driverType == "7") { currentSheet.Cell(8, 3).Value = "Raise/Lower Continuous"; }
                            if (driverType == "8") { currentSheet.Cell(8, 3).Value = "Multi Stage"; }
                        }
                        if (strategyData[4, i].Contains("I"))// get input label - data[11, i]=diginNumbers, data[12, i]=diginLabels
                        {
                            string inputNumber = strategyData[4, i].Substring(1, strategyData[4, i].Length - 1);
                            int n = 0;
                            while (strategyData[11, n] != inputNumber) { n++; }
                            currentSheet.Cell(8, 4).Value = strategyData[12, n];
                            currentSheet.Cell(8, 3).Value = "Digital Input";
                        }
                        //Soft Reference - need to sort out driver types
                        if (strategyData[4, i].Contains("L")) { currentSheet.Cell(8, 5).Value = strategyData[4, i].Remove(strategyData[4, i].Length - 1); } else { currentSheet.Cell(8, 5).Value = strategyData[4, i]; }
                        //Reference Type
                        currentSheet.Cell(8, 6).Value = strategyData[3, i];

                        rowIndex = 9;
                        worksheetIndex++;
                    }
                    else//Same sheet
                    {
                        //Channel
                        currentSheet.Cell(rowIndex, 1).Value = strategyData[1, i];
                        //Inputs/output
                        if (worksheetIndex == 2)
                        {
                            if (strategyData[3, i].Contains("I")) { currentSheet.Cell(rowIndex, 2).Value = "Inputs"; } else { currentSheet.Cell(rowIndex, 2).Value = "Outputs"; }
                        }
                        else
                        {
                            currentSheet.Cell(rowIndex, 2).Value = IORef;
                        }

                        //Soft Reference - need to sort out driver types
                        if (strategyData[4, i].Contains("L")) { currentSheet.Cell(rowIndex, 5).Value = strategyData[4, i].Remove(strategyData[4, i].Length - 1); } else { currentSheet.Cell(rowIndex, 5).Value = strategyData[4, i]; }
                        //Reference Type
                        currentSheet.Cell(rowIndex, 6).Value = strategyData[3, i];


                        //Labels and types
                        if (strategyData[4, i].Contains("S"))// get sensor label and type - data[8, i]=sensorNumber, data[9, i]=sensorLabel, data[10, i]=sensorType
                        {
                            string sensorNumber = strategyData[4, i].Substring(1, strategyData[4, i].Length - 1);
                            int s = 0;
                            while (strategyData[8, s] != sensorNumber) { s++; }
                            currentSheet.Cell(rowIndex, 4).Value = strategyData[9, s];
                            string SensorType = strategyData[10, s];//data[15, i]=typeNumbers, data[16, i]=setPartNumbers
                            int t = 0;
                            while (strategyData[15, t] != SensorType) { t++; }
                            currentSheet.Cell(rowIndex, 3).Value = strategyData[16, t];

                        }
                        if (strategyData[4, i].Contains("D"))// get driver label and type (digital or analogue, etc.) - data[13, i]=driverNumber, data[14, i]=driverLabel
                        {
                            string driverNumber = strategyData[4, i].Substring(1, strategyData[4, i].Length - 2);
                            int d = 0;
                            while (strategyData[13, d] != driverNumber) { d++; }
                            currentSheet.Cell(rowIndex, 4).Value = strategyData[14, d];
                            string driverType = strategyData[17, d];
                            if (driverType == "1") { currentSheet.Cell(rowIndex, 3).Value = "Digital Driver"; }
                            if (driverType == "2") { currentSheet.Cell(rowIndex, 3).Value = "Analogue Driver"; }
                            if (driverType == "3") { currentSheet.Cell(rowIndex, 3).Value = "Time Proportional"; }
                            if (driverType == "4") { currentSheet.Cell(rowIndex, 3).Value = "Raise/Lower End"; }
                            if (driverType == "5") { currentSheet.Cell(rowIndex, 3).Value = "Binary Histeresis"; }
                            if (driverType == "6") { currentSheet.Cell(rowIndex, 3).Value = "Time Proportional +O/R"; }
                            if (driverType == "7") { currentSheet.Cell(rowIndex, 3).Value = "Raise/Lower Continuous"; }
                            if (driverType == "8") { currentSheet.Cell(rowIndex, 3).Value = "Multi Stage"; }
                        }
                        if (strategyData[4, i].Contains("I"))// get input label - data[11, i]=diginNumbers, data[12, i]=diginLabels
                        {
                            string inputNumber = strategyData[4, i].Substring(1, strategyData[4, i].Length - 1);
                            int n = 0;
                            while (strategyData[11, n] != inputNumber) { n++; }
                            currentSheet.Cell(rowIndex, 4).Value = strategyData[12, n];
                            currentSheet.Cell(rowIndex, 3).Value = "Digital Input";
                        }
                        //setup cell options
                        currentSheet.Cell(8, 7).CopyTo(currentSheet.Cell(rowIndex, 7));
                        currentSheet.Cell(8, 8).CopyTo(currentSheet.Cell(rowIndex, 8));
                        currentSheet.Cell(8, 9).CopyTo(currentSheet.Cell(rowIndex, 9));
                        for (int st = 1; st < 12; st++)
                        {
                            currentSheet.Cell(rowIndex, st).Style = style1;
                        }

                        rowIndex++;


                    }
                    i++;
                    description = strategyData[2, i];
                    if (description != "") { IORef = description; IORef = strategyData[5, i] + " " + strategyData[18, i]; }
                }
                for (int d = 0; d < 32 - worksheetIndex; d++) { workbook.Worksheet(worksheetIndex).Delete(); }
            }
            catch(Exception myException)
            {
                int dummy = 0;
            }
        }


        public void SaveReport()
        {
            try
            {

                if (workbook.Worksheets.Count > 1)
                {
                    //workbook.Worksheet(1).Delete();
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Excel Workbook | *.xlsx";
                    Nullable<bool> result = saveFileDialog.ShowDialog();
                    if (result == true)
                    {
                        string fileName = System.IO.Path.GetFileName(saveFileDialog.FileName);
                        string fileDir = System.IO.Path.GetDirectoryName(saveFileDialog.FileName);
                        workbook.SaveAs(fileDir + "\\" + fileName);
                        MessageBox.Show("Excel document Data_Report.xlsx has been created succesfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Excel document not created - no data available!", "Warning", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception anException)
            {
                MessageBox.Show("Cannot access excel document, please make sure it is not open and directory is accessible.", "Warning", MessageBoxButton.OK, MessageBoxImage.Error);
                Console.WriteLine("Unexpected exception : {0}", anException.ToString() + " Cannot access excel document, please make sure it is not open and directory is accessible.");
            }
        }

        public bool AutoSaveReport(string lan, string os, string saveDirectory)
        {
            try
            {

                if (workbook.Worksheets.Count > 0)
                {
                    //string fileName2 = System.AppDomain.CurrentDomain.BaseDirectory + "Created PAC Sheets" + "\\" + siteName + "_L" + lan + "O" + os + "_PAC.xlsx";
                    string fileName2 = saveDirectory + siteName + "_L" + lan + "O" + os + "_PAC.xlsx";
                    workbook.SaveAs(fileName2);
                    return true;
                    //MessageBox.Show("Excel document Data_Report.xlsx has been created succesfully.");
                }
                else
                {
                    MessageBox.Show("Excel document not created - no data available!", "Warning", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }
            catch (Exception anException)
            {
                MessageBox.Show("Cannot access excel document, please make sure it is not open and directory is accessible.", "Warning", MessageBoxButton.OK, MessageBoxImage.Error);
                Console.WriteLine("Unexpected exception : {0}", anException.ToString() + " Cannot access excel document, please make sure it is not open and directory is accessible.");
                return false;
            }
        }



    }
}
