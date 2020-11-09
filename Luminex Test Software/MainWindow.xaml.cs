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

namespace Luminex_Test_Software
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 

    public partial class MainWindow : Window
    {
        private Dictionary<string, Label> labelDict;
        private readonly SerialController port;
        private readonly LuminexUnit sut; // System Under Test
        private readonly DocumentCreator documentCreator;

        public MainWindow()
        {
            InitializeComponent();
            InitLabelDict();
            port = new SerialController();
            sut = new LuminexUnit();
            documentCreator = new DocumentCreator();

            port.voltageDataDelegateCallback = VoltageDataReceived_Event;
            port.voltageReadStartedDelegateCallback = VoltageDataStarted_Event;
            port.voltageReadEndedDelegateCallback = VoltageDataEnded_Event;
            port.serialConnectedDelegateCallback = SerialConnected_Event;
            port.serialDisconnectedDelegateCallback = SerialDisconnected_Event;

            documentCreator.documentSavedDelegateCallback = DocumentSaveComplete_Event;
        }

        private void InitLabelDict()
        {
            labelDict = new Dictionary<string, Label>
            {
                { "_4140_1", _4140_1},
                { "_4140_2", _4140_2},
                { "_4140_3", _4140_3},
                { "_4140_4", _4140_4},
                { "_4140_5", _4140_5},

                { "_4130_1", _4130_1},
                { "_4130_2", _4130_2},
                { "_4130_3", _4130_3},
                { "_4130_4", _4130_4},
                { "_4130_5", _4130_5},
                { "_4130_7", _4130_7},
                { "_4130_8", _4130_8},
                { "_4130_9", _4130_9},
                { "_4130_10", _4130_10},
                { "_4130_11", _4130_11},
                { "_4130_12", _4130_12}
            };
        }

        public void DocumentSaveComplete_Event(bool success)
        {
            sut.ResetCables();
            textBoxSerialNumber.Clear();
            ResetLabels();
        }

        public void SerialConnected_Event()
        {
            Console.Out.WriteLine("Serial connected");
            Dispatcher.Invoke(new Action(() =>
            {
                label_Connection.Content = "Connected";
                button_ReadVoltage.IsEnabled = true;
            }));
        }

        public void SerialDisconnected_Event() { }

        public void VoltageDataStarted_Event()
        {
            ResetLabels();
        }

        public void VoltageDataEnded_Event()
        {
            //fill in for the 2 cable pins that provide common ground between respective cable and micro.  see schematic for details.
            //If all other pins are correct, then these ground pins have been validated.

            if(sut.pinDictionary["_4140_1"].passes() && sut.pinDictionary["_4140_3"].passes())
            {
                sut.pinDictionary["_4140_2"].MeasuredVoltage = 0.0;
                Dispatcher.Invoke(new Action(() => {
                    labelDict["_4140_2"].Content = $"{0}V";
                    labelDict["_4140_2"].Foreground = Brushes.Green;
                }));
            }

            if (sut.pinDictionary["_4130_2"].passes() && sut.pinDictionary["_4130_7"].passes())
            {
                sut.pinDictionary["_4130_1"].MeasuredVoltage = 0.0;
                Dispatcher.Invoke(new Action(() => {
                    labelDict["_4130_1"].Content = $"{0}V";
                    labelDict["_4130_1"].Foreground = Brushes.Green;
                }));
            }
        }

        public void VoltageDataReceived_Event(string cable, int pin, double voltage)
        {
            //Console.Out.WriteLine($"{cable}, pin: {pin}, voltage: {voltage}");
            string key = $"_{cable}_{pin}";
            sut.pinDictionary[key].MeasuredVoltage = voltage;

            Dispatcher.Invoke(new Action(() => { 
                labelDict[key].Content = $"{voltage}V"; 
                labelDict[key].Foreground = sut.pinDictionary[key].passes() ? Brushes.Green : Brushes.Red; 
            } ));
        }

        private void Button_ReadVoltage_Click(object sender, RoutedEventArgs e)
        {
            port.Read_Voltage();
        }

        private void Button_Save_Click(object sender, RoutedEventArgs e)
        {
            //new DocumentCreator().TestDocument();
            if (textBoxSerialNumber.Text.Equals(""))
            {
                MessageBox.Show("Enter Serial Number");
                return;
            }
            documentCreator.SaveResultsToFile(textBoxSerialNumber.Text, sut.pinDictionary);

        }

        private void ResetLabels()
        {
            Dispatcher.Invoke(new Action(() =>
            {
                foreach (KeyValuePair<string, Label> element in labelDict)
                {
                    element.Value.Content = "X";
                    element.Value.Foreground = Brushes.Black;
                }
            }));
        }
    }
}
