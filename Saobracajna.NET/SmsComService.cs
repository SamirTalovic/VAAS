using System;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SaobracajnaNET
{
    public class SmsComService
    {
        private SerialPort _port;

        public SmsComService(string comPort)
        {
            _port = new SerialPort(comPort, 115200, Parity.None, 8, StopBits.One)
            {
                Handshake = Handshake.None,
                Encoding = Encoding.ASCII,
                NewLine = "\r",
                ReadTimeout = 5000,
                WriteTimeout = 5000
            };
        }

        // ================= PORT =================

        public void Open()
        {
            if (_port.IsOpen)
                return;

            _port.Open();
            Thread.Sleep(2000);

            WakeUpModem();
            Send("ATE0"); // echo off
        }

        public void Close()
        {
            if (_port.IsOpen)
                _port.Close();
        }

        // ================= WAKE MODEM =================

        private void WakeUpModem()
        {
            for (int i = 0; i < 5; i++)
            {
                Send("AT");
                Thread.Sleep(500);

                string resp = _port.ReadExisting();
                if (resp.Contains("OK"))
                    break;
            }

            // signal
            Send("AT+CSQ");
            Thread.Sleep(500);
            _port.ReadExisting();

            // network registration
            Send("AT+CREG?");
            Thread.Sleep(500);
            _port.ReadExisting();

            // full functionality
            Send("AT+CFUN=1");
            Thread.Sleep(1000);

            // SMS text mode
            Send("AT+CMGF=1");
            Thread.Sleep(500);
        }

        // ================= USSD =================

        public async Task<string> SendUssdAsync(string ussd)
        {
            if (!_port.IsOpen)
                Open();

            WakeUpModem();

            _port.DiscardInBuffer();
            _port.DiscardOutBuffer();

            _port.Write($"AT+CUSD=1,\"{ussd}\",15\r");

            await Task.Delay(8000);

            string raw = _port.ReadExisting();
            return ParseUssd(raw);
        }

        public string ParseUssd(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return "PRAZAN ODGOVOR";

            int idx = raw.IndexOf("+CUSD:");
            if (idx < 0)
                return raw;

            int q1 = raw.IndexOf('"', idx);
            int q2 = raw.IndexOf('"', q1 + 1);

            if (q1 < 0 || q2 <= q1)
                return raw;

            string payload = raw.Substring(q1 + 1, q2 - q1 - 1);

            if (payload.Length % 2 == 0 && payload.All(Uri.IsHexDigit))
                return DecodeUcs2(payload);

            return payload;
        }

        private string DecodeUcs2(string hex)
        {
            byte[] bytes = new byte[hex.Length / 2];
            for (int i = 0; i < bytes.Length; i++)
                bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);

            return Encoding.BigEndianUnicode.GetString(bytes);
        }

        // ================= SMS =================

        public void SendSms(string phone, string message)
        {
            if (!_port.IsOpen)
                Open();

            WakeUpModem();

            phone = phone.Replace(" ", "").Replace("-", "");

            Send("AT+CMGF=1");
            Send($"AT+CMGS=\"{phone}\"");

            _port.Write(message + (char)26); // CTRL+Z
            Thread.Sleep(4000);
            _port.ReadExisting();
        }

        // ================= HELPERS =================

        private void Send(string cmd)
        {
            _port.Write(cmd + "\r");
            Thread.Sleep(400);
            _port.ReadExisting();
        }
    }
}
