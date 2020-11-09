using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Ports;
using System.Threading;
using System.Windows.Threading;

namespace Luminex_Test_Software
{

    public delegate void SerialConnected();
    public delegate void SerialDisconnected();
    public delegate void VoltageReadStarted();
    public delegate void VoltageReadEnded();
    public delegate void VoltageDataReceived(string cable, int pin, double voltage);

    public class SerialCommand
    {
        public SerialCommand(string value) { Value = value; }

        public string Value { get; set; }

        public static SerialCommand Handshake { get { return new SerialCommand("cmdHandshake"); } }
        public static SerialCommand ReadVoltage { get { return new SerialCommand("cmdReadVoltage"); } }
        public static SerialCommand PollConnection { get { return new SerialCommand("cmdConnectionPoll"); } }
    }

    class SerialController
    {

        private SerialPort arduino;
        private bool _connected = false;
        public bool Connected
        {
            get { return _connected; }
            set
            {
                _connected = value;
                if (_connected)
                {
                    serialConnectedDelegateCallback();
                    timer_AutoConnect.Stop();
                    //timer_ConnectionPoll.Start();
                } else
                {
                    InitArduino();
                    serialDisconnectedDelegateCallback();
                    timer_AutoConnect.Start();
                    //timer_ConnectionPoll.Stop();
                }
            }
        }
        private System.Timers.Timer timer_AutoConnect;
        //private System.Timers.Timer timer_ConnectionPoll;

        public VoltageReadStarted voltageReadStartedDelegateCallback;
        public VoltageReadEnded voltageReadEndedDelegateCallback;
        public VoltageDataReceived voltageDataDelegateCallback;
        public SerialConnected serialConnectedDelegateCallback;
        public SerialDisconnected serialDisconnectedDelegateCallback;

        public SerialController()
        {
            InitArduino();

            timer_AutoConnect = new System.Timers.Timer(2000);
            timer_AutoConnect.Elapsed += Timer_AutoConnect_Elapsed;
            timer_AutoConnect.AutoReset = true;
            timer_AutoConnect.Enabled = true;

            timer_AutoConnect.Start();

            //timer_ConnectionPoll = new System.Timers.Timer(500);
            //timer_ConnectionPoll.Elapsed += Timer_ConnectionPoll_Elapsed;
            //timer_ConnectionPoll.AutoReset = true;
            //timer_ConnectionPoll.Enabled = true;
        }

        private void InitArduino()
        {
            arduino = new SerialPort();
            arduino.DataReceived += Arduino_DataReceived;
            arduino.ErrorReceived += Arduino_ErrorReceived;
            arduino.Disposed += Arduino_Disposed;
            arduino.PinChanged += Arduino_PinChanged;
        }

        private void Timer_ConnectionPoll_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if(!Connected) { return; }
            try
            {
                Serial_Send_Command(SerialCommand.PollConnection);
            } catch (Exception error) {
                Console.Out.WriteLine($"Polling error.  Closing. {error}");
                Connected = false;
            }
        }

        private void Timer_AutoConnect_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if(!Connected)
            {
                PollPorts();
            } else
            {
                timer_AutoConnect.Stop();
            }
        }

        private void PollPorts()
        {

            string[] ports = SerialPort.GetPortNames();

            foreach (string port in ports)
            {

                if (Connected) { return; }

                if (arduino.IsOpen)
                {
                    arduino.Close();
                }

                Console.Out.WriteLine($"trying to connect to {port}");

                arduino.PortName = port;
                arduino.BaudRate = 9600;
                

                try
                {
                    //arduino.WriteLine("cmdHandshake");
                    arduino.Open();
                    Serial_Send_Command(SerialCommand.Handshake);
                }
                catch (Exception e)
                {
                    Console.Out.WriteLine(e.Message);
                    arduino.Close();
                }

                Task.Delay(2000).Wait();

            }
        }

        public void Read_Voltage()
        {
            try
            {
                Serial_Send_Command(SerialCommand.ReadVoltage);
            } catch (Exception e)
            {
                Console.Out.WriteLine($"Read Voltage Error: {e}");
                Connected = false;
            }
        }

        private void Serial_Send_Command(SerialCommand command)
        {
            try
            {
                arduino.Write(command.Value);
            }
            catch (System.IO.IOException e)
            {
                throw e;
            }
        }

        private void Arduino_PinChanged(object sender, SerialPinChangedEventArgs e)
        {
            Console.Out.WriteLine("Arduino pin changed." + e.ToString());
        }

        private void Arduino_Disposed(object sender, EventArgs e)
        {
            Console.Out.WriteLine("Arduino port disposed." + e.ToString());
        }

        private void Arduino_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            Console.Out.WriteLine("Arduino error received." + e.ToString());
        }

        private void Arduino_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string string_Received = arduino.ReadLine().TrimEnd('\r', '\n');
            Console.Out.WriteLine(string_Received);

            if (string_Received.Equals("Connected")) {
                Connected = true;
            } 
            else if (string_Received.Equals("voltageReadStart")) {
                voltageReadStartedDelegateCallback();
            } 
            else if (string_Received.Equals("voltageReadEnd")) {
                voltageReadEndedDelegateCallback();
            } else {
                string[] data = string_Received.Split(',');
                int pin = int.Parse(data[1]);
                double voltage = double.Parse(data[2]);
                voltageDataDelegateCallback(data[0], pin, voltage);
            }

            //voltageDataDelegateCallback("lskdjf", 4, 2.39420);
        }
    }
}
