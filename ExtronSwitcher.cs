// C# Extron A/V Switcher Control Class
//
// Provides a clean interface to all Extron switcher gear supporting the same
// protocol as the SW-6AV and SW-12AV switchers.  Most of this class was
// developed against the serial protocol specification in section 4 of the
// SW-AV series user manual.
//
// Licensed under the MIT license:
//
// Copyright (c) 2004 Andrew Paprocki <andrew@ishiboo.com>
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy 
// of this software and associated documentation files (the "Software"), to 
// deal in the Software without restriction, including without limitation the 
// rights to use, copy, modify, merge, publish, distribute, sublicense, and/or 
// sell copies of the Software, and to permit persons to whom the Software is 
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in 
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS 
// IN THE SOFTWARE.

using System;
using System.Text;

using LoMaN.IO;

namespace Ishiboo
{
    public class ExtronSwitcher : IDisposable
    {
        private SerialStream port = null;
        private byte[] portstatus = null;
        private bool disposed = false;

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (this.disposed)
                return;

            if (disposing)
            {
                this.port.Close();
                this.port = null;

                this.portstatus = null;
            }

            this.disposed = true;
        }

        public enum SwitchModeType
        {
            Normal = 0,
            Auto = 1
        }

        public int Ports
        {
            get { return this.portstatus.Length; }
        }

        private uint selectedaudioport = 1;
        public uint SelectedAudioPort
        {
            get { return this.selectedaudioport; }
        }

        private uint selectedvideoport = 1;
        public uint SelectedVideoPort
        {
            get { return this.selectedvideoport; }
        }

        public SwitchModeType SwitchMode 
        {
            get
            {
                byte[] b = new byte[1];
                b[0] = Convert.ToByte('I');

                this.port.Write(b, 0, b.Length);
                this.port.Flush();

                string r = this.ParseInput();
                if (!r.StartsWith("V*"))
                    throw new Exception("Invalid panel mode response '" + r + "'");

                string[] s = r.Split(new char[] {' '});
                if (s.Length != 4)
                    throw new Exception("Invalid panel mode response '" + r + "'");

                uint sm = Convert.ToUInt32(s[2].Substring(2, 1));

                if (sm != 1 && sm != 2)
                    throw new Exception("Invalid panel mode response '" + r + "'");

                if (sm == 1)
                    return SwitchModeType.Normal;
                else
                    return SwitchModeType.Auto;
            }
            set
            {
                byte[] b = new byte[2];
                b[0] = (value == SwitchModeType.Normal) ? Convert.ToByte('1') : Convert.ToByte('2');
                b[1] = Convert.ToByte('#');

                this.port.Write(b, 0, b.Length);
                this.port.Flush();

                string r = this.ParseInput();
                if (r.Length != 2 || !r.StartsWith("F"))
                    throw new Exception("Invalid panel mode response '" + r + "'");

                uint sm = Convert.ToUInt32(r.Substring(1));

                if ((sm != 1 && sm != 2) ||
                    (sm == 1 && value != SwitchModeType.Normal) ||
                    (sm == 2 && value != SwitchModeType.Auto))
                    throw new Exception("Invalid panel mode response '" + r + "'");
            }
        }

        public bool VideoMuted
        {
            get
            {
                byte[] b = new byte[1];
                b[0] = Convert.ToByte('B');

                this.port.Write(b, 0, b.Length);
                this.port.Flush();

                string r = this.ParseInput();
                if (r.Length != 1)
                    throw new Exception("Invalid muted response '" + r + "'");

                uint s = Convert.ToUInt32(r);

                if (s != 0 && s != 1)
                    throw new Exception("Invalid muted response '" + r + "'");

                if (s == 0)
                    return false;
                else
                    return true;
            }
            set
            {
                byte[] b = new byte[2];
                b[0] = value ? Convert.ToByte('1') : Convert.ToByte('0');
                b[1] = Convert.ToByte('B');

                this.port.Write(b, 0, b.Length);
                this.port.Flush();

                string r = this.ParseInput();
                if (!r.StartsWith("Vmt"))
                    throw new Exception("Invalid muted response '" + r + "'");

                if (Convert.ToBoolean(r.Substring(3)) != value)
                    throw new Exception("Invalid muted status '" + r + "'");
            }
        }

        public bool AudioMuted
        {
            get
            {
                byte[] b = new byte[1];
                b[0] = Convert.ToByte('Z');

                this.port.Write(b, 0, b.Length);
                this.port.Flush();

                string r = this.ParseInput();
                if (r.Length != 1)
                    throw new Exception("Invalid muted response '" + r + "'");

                return Convert.ToBoolean(r);
            }
            set
            {
                byte[] b = new byte[2];
                b[0] = value ? Convert.ToByte('1') : Convert.ToByte('0');
                b[1] = Convert.ToByte('Z');

                this.port.Write(b, 0, b.Length);
                this.port.Flush();

                string r = this.ParseInput();
                if (!r.StartsWith("Amt"))
                    throw new Exception("Invalid muted response '" + r + "'");

                if (Convert.ToBoolean(r.Substring(3)) != value)
                    throw new Exception("Invalid muted status '" + r + "'");
            }
        }

        public string PartNumber
        {
            get
            {
                byte[] b = new byte[1];
                b[0] = Convert.ToByte('N');

                this.port.Write(b, 0, b.Length);
                this.port.Flush();

                string r = this.ParseInput();
                if (r == null)
                    throw new Exception("Invalid version response");

                return r;
            }
        }

        public string Version
        {
            get
            {
                byte[] b = new byte[1];
                b[0] = Convert.ToByte('Q');

                this.port.Write(b, 0, b.Length);
                this.port.Flush();

                string r = this.ParseInput();
                if (r == null)
                    throw new Exception("Invalid version response");

                return r;
            }
        }

        public ExtronSwitcher(uint comport)
        {
            if (comport == 0)
                throw new ArgumentOutOfRangeException();

            this.port = new SerialStream("COM" + comport.ToString());

            this.port.SetPortSettings(9600, SerialStream.FlowControl.None,
                                      SerialStream.Parity.None, 8,
                                      SerialStream.StopBits.One);
            this.port.SetTimeouts(0, 0, 500, 0, 500);

            this.RequestInputStatus();
        }

        private string ParseInput()
        {
            ASCIIEncoding ae = new ASCIIEncoding();
            byte[] b = new byte[1024];
            int i = this.port.Read(b);
            string bs = ae.GetString(b);
            bs = bs.Replace("\r", "");
            bs = bs.Replace("\0", "");
            string[] s = bs.Split(new char[] {'\n'});

            foreach (string str in s)
            {
                if (str.StartsWith("(C)"))
                {// copyright message
                    Console.WriteLine("Extron: Reboot: " + str);
                }
                else if (str.StartsWith("E"))
                {// error
                    uint etype = Convert.ToUInt32(str.Substring(1));
                    if (etype == 1)
                        throw new Exception("Invalid input channel number (out of range)");
                    else if (etype == 10)
                        throw new Exception("Invalid command");
                    else if (etype == 13)
                        throw new Exception("Invalid parameter (out of range)");
                    else if (etype == 14)
                        throw new Exception("Illegal command for this configuration");
                    else
                        throw new Exception("Error '" + str.Substring(1) + "' received");
                }
                else if (str.StartsWith("C"))
                {// selected port changed
                    uint port = Convert.ToUInt32(str.Substring(1));

                    if (port == 0 || port > this.portstatus.Length)
                        throw new Exception("Invalid 'C' selected port " + port);

                    this.selectedaudioport = port;
                    this.selectedvideoport = port;
                }
                else if (str.StartsWith("Sig"))
                {// input status message
                    string[] status = str.Split(new char[] {' '});
                    if (status.Length <= 0)
                        throw new Exception("Invalid 'Sig' message" + str);
                    this.portstatus = new byte[status.Length];
                    for (int j = 1; j < status.Length; j++)
                    {
                        if (status[j] == "0")
                            this.portstatus[j - 1] = 0;
                        else if (status[j] == "1")
                            this.portstatus[j - 1] = 1;
                        else
                            throw new Exception("Invalid 'Sig' status on port " + j);

                        Console.WriteLine("Extron: Port({0}) Status({1})", j, status);
                    }
                }
                else if (str.StartsWith("In"))
                {
                    string[] status = str.Split(new char[] {' '});
                    if (status.Length != 2)
                        throw new Exception("Invalid 'In' message: " + str);

                    uint port = Convert.ToUInt32(status[0].Substring(2));
                    if (port > this.portstatus.Length)
                        throw new Exception("Invalid input " + port + " specified");

                    if (status[1] == "All")
                    {
                        this.selectedaudioport = port;
                        this.selectedvideoport = port;
                    }
                    else if (status[1] == "Aud")
                        this.selectedaudioport = port;
                    else if (status[1] == "Vid")
                        this.selectedvideoport = port;
                    else
                        throw new Exception("Invalid selection type '" + status[1] + "'");

                    Console.WriteLine("Extron: Port({0}) Input({1}) selected", port, status[1]);
                }
                else if (str.StartsWith("Reconfig"))
                {// audio automatic gain readjustment
                    Console.WriteLine("Extron: Reconfig initiated");
                }
                else
                {// command specific response message
                    return str;
                }
            }

            return null;
        }

        public void RequestInputStatus()
        {
            byte[] b = new byte[2];
            b[0] = Convert.ToByte('0');
            b[1] = Convert.ToByte('S');

            this.port.Write(b, 0, b.Length);
            this.port.Flush();

            this.ParseInput();
        }

        public bool RequestInputStatus(uint input)
        {
            if (input == 0 || input > this.portstatus.Length)
                throw new ArgumentOutOfRangeException();

            byte[] b = new byte[2];
            b[0] = Convert.ToByte(input.ToString()[0]); 
            b[1] = Convert.ToByte('S');

            this.port.Write(b, 0, b.Length);
            this.port.Flush();

            string r = this.ParseInput();

            return Convert.ToBoolean(r);
        }

        public void SelectInputAudio(uint input)
        {
            if (input == 0 || input > this.portstatus.Length)
                throw new ArgumentOutOfRangeException();

            byte[] b = new byte[2];
            b[0] = Convert.ToByte(input.ToString()[0]);
            b[1] = Convert.ToByte('$');

            this.port.Write(b, 0, b.Length);
            this.port.Flush();
        }

        public void SelectInputVideo(uint input)
        {
            if (input == 0 || input > this.portstatus.Length)
                throw new ArgumentOutOfRangeException();

            byte[] b = new byte[2];
            b[0] = Convert.ToByte(input.ToString()[0]);
            b[1] = Convert.ToByte('&');

            this.port.Write(b, 0, b.Length);
            this.port.Flush();
        }

        public void SelectInputAudioVideo(uint input)
        {
            if (input == 0 || input > this.portstatus.Length)
                throw new ArgumentOutOfRangeException();

            byte[] b = new byte[2];
            b[0] = Convert.ToByte(input.ToString()[0]);
            b[1] = Convert.ToByte('!');

            this.port.Write(b, 0, b.Length);
            this.port.Flush();
        }
    }
}
