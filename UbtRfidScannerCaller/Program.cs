using System;
using UbtRfidScanner;

namespace UbtRfidScannerCaller
{
    class Program
    {
        public delegate void eReadyHandler(string rdif);

        static void Main(string[] args)
        {
            //Type t = Type.GetTypeFromProgID("UbtRfidScanner.ComComponent");
            //dynamic component = Activator.CreateInstance(t);
            var component = new ComComponent();

            bool result = component.ConnectComPort(5);

            component.OnRfidArrived2 += new EventHandler<string>(OnRfidArrived2);

            if (!component.ConnectComPort(5))
            {
                Console.WriteLine("Не удалось подключиться " + component.LastReturnCodeDescription());
            }
            else
            {
                Console.WriteLine("Подключились");
            }

            //component.OnRfidArrived += OnRfidArrived;
            component.StartScanning();

            Console.ReadKey();
            //component.DisconnectComPort();
        }

        private static void OnRfidArrived(string rfid)
        {
            Console.WriteLine("RFID arrived " + rfid);
        }

        private static void OnRfidArrived2(object sender, string rfid)
        {
            Console.Write("RFID 2 arrived " + rfid);
        }
    }
}
