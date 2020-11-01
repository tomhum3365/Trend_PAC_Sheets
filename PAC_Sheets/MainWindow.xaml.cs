using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.OleDb;
using System.Windows.Threading;

//Author:   Tomas Humplik
//Date:     24.10.2020

namespace PAC_Sheets
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            createButton.IsEnabled = false;
            Loaded += (sender, e) =>MoveFocus(new TraversalRequest(FocusNavigationDirection.First));
        }

        //String readQuery = "SELECT Meter FROM MeterIDs ORDER BY ID ASC";
        string identifierQuery = "SELECT identifier FROM IQAddressModule";
        //string lanQuery = "SELECT local_lan FROM IQAddressModule";
        string lanQuery = "SELECT lan FROM DeviceDetails";
        //string osQuery = "SELECT local_node FROM IQAddressModule";
        string osQuery = "SELECT node FROM DeviceDetails";
        string versionQuery = "SELECT device_name FROM DeviceDetails";
        string channelNumberQuery = "SELECT channel_number FROM IOAssignment";
        string descriptionQuery = "SELECT description FROM IOAssignment";
        string channelTypeQuery = "SELECT channel_type FROM IOAssignment";
        string channelAssignmentQuery = "SELECT channel_assignment FROM IOAssignment";
        string idNumberQuery = "SELECT id_number FROM IOAssignment";
        string valueQuery = "SELECT value_1 FROM IOAssignment";
        string serialQuery = "SELECT serial_number FROM IOAssignment";
        string sensorNumberQuery = "SELECT module_number FROM IQSensorType1";
        string sensorLabelQuery = "SELECT label FROM IQSensorType1";
        string sensorTypeQuery = "SELECT sensor_type FROM IQSensorType1";
        string diginNumberQuery = "SELECT module_number FROM IQDiginType1";
        string diginLabelQuery = "SELECT label FROM IQDiginType1";
        string driverNumberQuery = "SELECT module_number FROM IQDriver";
        string driverLabelQuery = "SELECT details FROM IQDriver";
        string typeNumberQuery = "SELECT type_number FROM IQSensorTypes";
        string setPartNumberQuery = "SELECT set_part_number FROM IQSensorTypes";
        string driverTypeQuery = "SELECT type FROM IQDriver";
        string ioTypeQuery = "SELECT io_type FROM IOAssignment";

        string aLan = "";
        string aOs = "";
        string[] directories;

        string pacDirectory = "";

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            //first reset the directories!!!
            directoriesTextBox.Text = "Selected files: \n";

            if (directories != null) { Array.Clear(directories, 0, directories.Length); }

            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
            openFileDlg.Filter = "Strategy files (*.IQ)|*.IQ|" + "All files (*.*)|*.*";

            openFileDlg.Multiselect = true;
            // Launch OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = openFileDlg.ShowDialog();
            // Get the selected file name and display in a TextBox.
            // Load content of file in a TextBlock
            if (result == true)
            {
                //directoryTextBox.Text = openFileDlg.FileName;
                directories = openFileDlg.FileNames;
                //get current directory (not full file directory, just folder)
                pacDirectory = directories[0].Substring(0, directories[0].LastIndexOf("\\") + 1) + "PAC Sheets\\";
                //TextBlock1.Text = System.IO.File.ReadAllText(openFileDlg.FileName);
                for (int d = 0; d < directories.Length; d++)
                {
                    directoriesTextBox.Text = directoriesTextBox.Text + directories[d] + "\n";
                    MyScrollViewer.ScrollToBottom();
                    Dispatcher.CurrentDispatcher.Invoke(DispatcherPriority.Background, (Action)(() => { }));
                }
                createButton.IsEnabled = true;
            }
            else
            {
                createButton.IsEnabled = false;
            }
        }

        private void createButton_Click(object sender, RoutedEventArgs e)
        {
            directoriesTextBox.Text = directoriesTextBox.Text + "\nCreating PAC sheets: \n";
            bool created = false;
            //loop for the directories length   for(
            for (int i = 0; i < directories.Length; i++)
            {
                try
                {
                    //Get data from strategy
                    string[,] strategyData = GetData(directories[i]);// if exception then catch and inform user that file is not valid IQ4 file

                    if (strategyData[3, 0] == null)
                    {
                        directoriesTextBox.Text = directoriesTextBox.Text + "WARNING - " + directories[i] + " Controller does not have any inputs or outputs!" + "\n";
                    }
                    else
                    {

                        // Create excel spreadsheets
                        ExcelDataManager ExcelManager = new ExcelDataManager();
                        ExcelManager.InsertData(strategyData);
                        //ExcelManager.SaveReport();
                        created = ExcelManager.AutoSaveReport(aLan, aOs, pacDirectory);//without directory and file name selector
                        if (created)
                        {
                            string siteName = siteTextBox.Text;
                            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                            {
                                siteName = siteName.Replace(c, '-');
                            }
                            //directoriesTextBox.Text = directoriesTextBox.Text + "Created: " + System.AppDomain.CurrentDomain.BaseDirectory + "PAC Sheets" + "\\" + siteTextBox.Text + "_L" + aLan + "O" + aOs + "_PAC.xlsx." + "\n";
                            directoriesTextBox.Text = directoriesTextBox.Text + "Created: " + pacDirectory + siteName + "_L" + aLan + "O" + aOs + "_PAC.xlsx." + "\n";
                            MyScrollViewer.ScrollToBottom();
                            Dispatcher.CurrentDispatcher.Invoke(DispatcherPriority.Background, (Action)(() => { }));
                        }
                        else
                        {
                            directoriesTextBox.Text = directoriesTextBox.Text + "Error - file not created! Please close Excel app.";
                        }
                        int dummy = 0;
                    }
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("IOAssignment"))
                    {
                        directoriesTextBox.Text = directoriesTextBox.Text + "WARNING - File: " + directories[i] + " is not a valid IQ4 strategy file!" + "\n";
                        MyScrollViewer.ScrollToBottom();
                        Dispatcher.CurrentDispatcher.Invoke(DispatcherPriority.Background, (Action)(() => { }));
                    }
                    int test = 1;
                }
            }
            if (created)
            {
                //MessageBox.Show("Excel documents have been created succesfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information); 
                directoriesTextBox.Text = directoriesTextBox.Text + "\nFinished creating the files.";
            }
        }

        public static void ForceUIToUpdate()
        {
            DispatcherFrame frame = new DispatcherFrame();

            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Render, new DispatcherOperationCallback(delegate (object parameter)
            {
                frame.Continue = false;
                return null;
            }), null);

            Dispatcher.PushFrame(frame);
        }
        private string[,] GetData(string directory)
        {
            string name = nameTextBox.Text;
            string date = DateTime.Now.ToString("dddd , MMM dd yyyy,hh:mm:ss");
            string site = siteTextBox.Text;
            // Get all required data from strategy
            AccessDBManager AccessManager = new AccessDBManager();
            AccessManager.SetDBlocation(directory);
            //IQAddressModule table
            string identifier = AccessManager.ReadData(identifierQuery)[0];
            string lan = AccessManager.ReadData(lanQuery)[0];
            string os = AccessManager.ReadData(osQuery)[0];
            //DeviceDetails
            string version = AccessManager.ReadData(versionQuery)[0];
            //IOAssignment table
            string[] channels = AccessManager.ReadData(channelNumberQuery);
            string[] description = AccessManager.ReadData(descriptionQuery);
            string[] channelType = AccessManager.ReadData(channelTypeQuery);
            string[] channelAssignment = AccessManager.ReadData(channelAssignmentQuery);
            string[] idNumber = AccessManager.ReadData(idNumberQuery);
            string[] value = AccessManager.ReadData(valueQuery);
            string[] serial = AccessManager.ReadData(serialQuery);
            string[] ioType = AccessManager.ReadData(ioTypeQuery);
            //IQSensorType1 table
            string[] sensorNumbers = AccessManager.ReadData(sensorNumberQuery);
            string[] sensorLabels = AccessManager.ReadData(sensorLabelQuery);
            string[] sensorTypes = AccessManager.ReadData(sensorTypeQuery);
            //IQDiginType1 table
            string[] diginNumbers = AccessManager.ReadData(diginNumberQuery);
            string[] diginLabels = AccessManager.ReadData(diginLabelQuery);
            //IQDriver table
            string[] driverNumbers = AccessManager.ReadData(driverNumberQuery);
            string[] driverLabels = AccessManager.ReadData(driverLabelQuery);
            string[] driverTypes = AccessManager.ReadData(driverTypeQuery);
            //IQSensorTypes table
            string[] typeNumbers = AccessManager.ReadData(typeNumberQuery);
            string[] setPartNumbers = AccessManager.ReadData(setPartNumberQuery);

            string[,] data = new string[19, 1000];
            for (int i = 0; i < 1000; i++)
            {
                data[0, 0] = name;
                data[0, 1] = date;
                data[0, 2] = identifier;
                data[0, 3] = lan;
                data[0, 4] = os;
                data[0, 5] = site;
                data[0, 6] = version;

                data[1, i] = channels[i];
                data[2, i] = description[i];
                data[3, i] = channelType[i];
                data[4, i] = channelAssignment[i];
                data[5, i] = idNumber[i];
                data[6, i] = value[i];
                data[7, i] = serial[i];
                data[8, i] = sensorNumbers[i];
                data[9, i] = sensorLabels[i];
                data[10, i] = sensorTypes[i];
                data[11, i] = diginNumbers[i];
                data[12, i] = diginLabels[i];
                data[13, i] = driverNumbers[i];
                data[14, i] = driverLabels[i];
                data[15, i] = typeNumbers[i]; //sensor types
                data[16, i] = setPartNumbers[i];//sensor types
                data[17, i] = driverTypes[i];
                data[18, i] = ioType[i];

                aLan = lan;
                aOs = os;
            }
            return data;
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("PAC Sheets v1.29" + Environment.NewLine + "Author: Tomas Humplik" + Environment.NewLine + "Date: 27/10/2020" + Environment.NewLine + Environment.NewLine + "Please report issues and new ideas to tomhum3365@gmail.com." + Environment.NewLine + Environment.NewLine + 
                "If you realy like this app please make a donation by going to Help->Donate." + Environment.NewLine + "Thank you", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void InstructionsItem_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://tomashumplik.com/PACSheets/publish.htm");
        }
        private void DonateItem_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.paypal.com/donate?hosted_button_id=DX6XXAXZETXRY");
        }
        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }
        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("For future use.", "Settings", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
