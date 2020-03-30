using System;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

/// <summary>
/// Отрефакторимся чуть позже
/// </summary>
namespace UbtRfidScanner
{
    public enum ConnectionState
    {
        Undefined, Successful, Error
    }

    [ComVisible(true)]
    [Guid("1C12CE52-3D59-4050-9F36-4192A60654E0")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRfidScannerEvents
    {
        [DispId(1)]
        void OnRfidArrived(string rfid);

        [DispId(2)]
        void OnRfidArrived2(object sender, string rfid);
    }

    [ComVisible(true)]
    [Guid("A44135B9-B1FB-4A27-B7B6-8AC803C2F655")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRfidScannerComponent
    {
        [DispId(1)]
        bool StartScanning();

        [DispId(2)]
        bool ConnectComPort(int comNumber);

        [DispId(3)]
        bool DisconnectComPort();

        [DispId(4)]
        string LastReturnCodeDescription();

        [DispId(5)]
        ConnectionState ConnectionState { get; }

        [DispId(6)]
        void RemoveAllEvents();

        [DispId(7)]
        void StopScanning();
    }

    [ComVisible(true)]
    [Guid("BC9FB3A4-17B6-4366-B004-8A535A630386")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IRfidScannerEvents))]
#if DEBUG
    public class ComComponent: IRfidScannerComponent
#else
    public class ComComponent:  ServicedComponent, IRfidScannerComponent
#endif
    {
        [ComVisible(false)]
        public delegate void OnRfidArrivedDelegate(string rfid);
        public event OnRfidArrivedDelegate OnRfidArrived;

        public event EventHandler<string> OnRfidArrived2;

        private int _comNumber = -1;
        private byte _comAddress = 255;
        private int _commandReturnCode = 30;
        public ConnectionState ConnectionState { get; private set; } = ConnectionState.Undefined;
        private CancellationTokenSource _scanCancellationTokenSource;

        public string LastReturnCodeDescription()
        {
            return _commandReturnCode.ToReturnCodeDescription();
        }

        public void RemoveAllEvents()
        {
            OnRfidArrived = null;
            OnRfidArrived2 = null;
        }

        public bool ConnectComPort(int comNumber)
        {

            OnRfidArrived?.Invoke("Коннекчусь пытаюсь");
            OnRfidArrived2?.Invoke(this, "Коннекчусь по второму методу");

            _comNumber = comNumber;
            _comAddress = 255;
            int frmPortIndex = 0;

            DisconnectComPort();
            _commandReturnCode = DllInteropBridge.OpenComPort(_comNumber, ref _comAddress, 0, ref frmPortIndex);

            if (_commandReturnCode != 0)
            {
                return false;
            }

            ConnectionState = ConnectionState.Successful;
            return true;
        }

        public bool DisconnectComPort()
        {
            StopScanning();
            _commandReturnCode = DllInteropBridge.CloseSpecComPort(_comNumber);

            if (_commandReturnCode != 0)
            {
                ConnectionState = ConnectionState.Error;
                return false;
            }

            return true;
        }

        public bool StartScanning()
        {
            if (ConnectionState != ConnectionState.Successful)
            {
                return false;
            }

            StopScanning();

            
            var token  = _scanCancellationTokenSource.Token;
            var task = ScheduleScanningAsyncRecursively(token);
            return true;
        }

        public void StopScanning()
        {
            _scanCancellationTokenSource?.Cancel();
            _scanCancellationTokenSource?.Dispose();

            _scanCancellationTokenSource = new CancellationTokenSource();
        }

        private Task ScheduleScanningAsyncRecursively(CancellationToken token)
        {
            if (token.IsCancellationRequested)
            {
                return Task.FromResult(true);
            }
            return Task.Run(() => RunScanning(), token)
                .ContinueWith(t => ScheduleScanningAsyncRecursively(token));
        }

        /// <summary>
        /// Поверхностный рефакторинг китайского кода.
        /// В основном оставил без изменений.
        /// </summary>
        /// <returns></returns>
        private void RunScanning()
        {
            byte scanTime = 10;
            byte qValue = 0;
            byte session = 0;
            byte tidFlag = 0; // !_rbEpc.Checked в примере от китайцев
            byte target = 0;

            byte[] epc = new byte[50000];
            byte maskMem = 0;
            byte[] maskAdr = new byte[2];
            byte maskLen = 0;
            byte[] maskData = new byte[100];
            byte maskFlag = 0;
            byte adrTID = 0;
            byte lenTID = 6;

            byte inAnt = 0x80;
            byte fastFlag = 1;

            byte ant = 0;
            int totalLen = 0;
            int cardNum = 0;

            _commandReturnCode = DllInteropBridge.Inventory_G2(
                ref _comAddress,
                qValue,
                session,
                maskMem,
                maskAdr,
                maskLen,
                maskData,
                maskFlag,
                adrTID,
                lenTID,
                tidFlag,
                target,
                inAnt,
                scanTime,
                fastFlag,
                epc,
                ref ant,
                ref totalLen,
                ref cardNum,
                _comNumber
                );

            if ((_commandReturnCode == 1)  | (_commandReturnCode == 2) | (_commandReturnCode == 3) | (_commandReturnCode == 4))
            {
                byte[] daw = new byte[totalLen];
                Array.Copy(epc, daw, totalLen);
                var temps = daw.ToHexString();
                var m = 0;

                for (int cardIndex = 0; cardIndex < cardNum; cardIndex++)
                {
                    var epcLen = daw[m] + 1;
                    var temp = temps.Substring(m * 2 + 2, epcLen * 2);
                    var sEPC = temp.Substring(0, temp.Length - 2);
                    var RSSI = Convert.ToInt32(temp.Substring(temp.Length - 2, 2), 16);

                    m += epcLen + 1;

                    if (sEPC.Length != (epcLen - 1) * 2)
                    {
                        // todo сделать ошибку отдельным событием.
                        throw new Exception("Не знаю пока что такое");
                    }

                    OnRfidArrived?.Invoke(sEPC);
                    OnRfidArrived2?.Invoke(this, sEPC);
                }
            }            
        }
    }

    public class RfidArrivedEventArgs : EventArgs
    {
        public readonly string Rfid;

        public RfidArrivedEventArgs(string rfid)
        {
            Rfid = rfid;
        }
    }
}
