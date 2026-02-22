using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace WhisperDictation
{
    static class Program
    {
        // --- Config ---
        static string GROQ_API_KEY;
        const string WHISPER_MODEL = "whisper-large-v3-turbo";
        const int HOTKEY_VK = 0x91; // VK_SCROLL
        const int SAMPLE_RATE = 16000;
        const int BITS = 16;
        const int CHANNELS = 1;
        const int MAX_RECORD_SEC = 60;

        static string exeDir;
        static string wavPath;
        static string logPath;

        // --- Native imports ---
        #region waveIn
        [StructLayout(LayoutKind.Sequential)]
        struct WAVEFORMATEX
        {
            public ushort wFormatTag, nChannels;
            public uint nSamplesPerSec, nAvgBytesPerSec;
            public ushort nBlockAlign, wBitsPerSample, cbSize;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct WAVEHDR
        {
            public IntPtr lpData;
            public uint dwBufferLength, dwBytesRecorded;
            public IntPtr dwUser;
            public uint dwFlags, dwLoops;
            public IntPtr lpNext, reserved;
        }

        [DllImport("winmm.dll")] static extern int waveInGetNumDevs();
        [DllImport("winmm.dll")] static extern int waveInOpen(out IntPtr phwi, int id, ref WAVEFORMATEX fmt, IntPtr cb, IntPtr inst, int flags);
        [DllImport("winmm.dll")] static extern int waveInPrepareHeader(IntPtr hwi, IntPtr pwh, int sz);
        [DllImport("winmm.dll")] static extern int waveInUnprepareHeader(IntPtr hwi, IntPtr pwh, int sz);
        [DllImport("winmm.dll")] static extern int waveInAddBuffer(IntPtr hwi, IntPtr pwh, int sz);
        [DllImport("winmm.dll")] static extern int waveInStart(IntPtr hwi);
        [DllImport("winmm.dll")] static extern int waveInStop(IntPtr hwi);
        [DllImport("winmm.dll")] static extern int waveInReset(IntPtr hwi);
        [DllImport("winmm.dll")] static extern int waveInClose(IntPtr hwi);
        #endregion

        #region Keyboard Hook
        const int WH_KEYBOARD_LL = 13;
        const int WM_KEYDOWN = 0x100, WM_KEYUP = 0x101, WM_SYSKEYDOWN = 0x104, WM_SYSKEYUP = 0x105;

        [StructLayout(LayoutKind.Sequential)]
        struct KBDLLHOOKSTRUCT { public int vkCode, scanCode, flags, time; public IntPtr dwExtraInfo; }

        delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")] static extern IntPtr SetWindowsHookEx(int id, LowLevelKeyboardProc proc, IntPtr hMod, uint tid);
        [DllImport("user32.dll")] static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")] static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")] static extern IntPtr GetModuleHandle(string name);
        #endregion

        #region SendInput
        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT { public int dx, dy; public uint mouseData, dwFlags, time; public IntPtr dwExtraInfo; }
        [StructLayout(LayoutKind.Sequential)]
        struct KEYBDINPUT { public ushort wVk, wScan; public uint dwFlags, time; public IntPtr dwExtraInfo; }

        [StructLayout(LayoutKind.Explicit)]
        struct INPUTUNION
        {
            [FieldOffset(0)] public MOUSEINPUT mi;  // largest member, sets union size
            [FieldOffset(0)] public KEYBDINPUT ki;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct INPUT { public uint type; public INPUTUNION u; }

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint SendInput(uint n, INPUT[] inputs, int size);

        #endregion

        [DllImport("kernel32.dll")]
        static extern bool Beep(uint freq, uint duration);

        // --- State ---
        static IntPtr hookId;
        static LowLevelKeyboardProc hookProc; // prevent GC
        static bool isRecording;
        static volatile bool startFlag, stopFlag;

        // waveIn state
        static IntPtr hwi, hdrPtr, bufPtr;
        static readonly int bufSize = SAMPLE_RATE * (BITS / 8) * CHANNELS * MAX_RECORD_SEC;

        [STAThread]
        static void Main(string[] args)
        {
            // TLS 1.2 must be set before any HTTP calls
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            GROQ_API_KEY = Environment.GetEnvironmentVariable("GROQ_API_KEY");
            if (string.IsNullOrEmpty(GROQ_API_KEY))
            {
                Console.WriteLine("ERROR: GROQ_API_KEY environment variable not set.");
                return;
            }

            exeDir = AppDomain.CurrentDomain.BaseDirectory;
            wavPath = Path.Combine(exeDir, "dictation_recording.wav");
            logPath = Path.Combine(exeDir, "dictation.log");

            File.WriteAllText(logPath, "");
            Log("=== WhisperDictation starting ===");

            // Install hook
            hookProc = HookCallback;
            using (var proc = Process.GetCurrentProcess())
            using (var mod = proc.MainModule)
                hookId = SetWindowsHookEx(WH_KEYBOARD_LL, hookProc, GetModuleHandle(mod.ModuleName), 0);

            if (hookId == IntPtr.Zero) { Log("ERROR: Hook install failed"); return; }
            Log("Hook installed");

            // Timer polls flags
            var timer = new System.Windows.Forms.Timer();
            timer.Interval = 30;
            timer.Tick += OnTimerTick;
            timer.Start();

            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine(" Push-to-Talk Dictation (Groq Whisper)");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("Hotkey:  Scroll Lock (hold to record)");
            Console.WriteLine("Model:   " + WHISPER_MODEL);
            Console.WriteLine("Log:     " + logPath);
            Console.WriteLine();
            Console.WriteLine("High beep = recording | Low beep = transcribing");
            Console.WriteLine("Close window or Ctrl+C to exit.");
            Console.WriteLine();
            Log("Listening...");

            Application.Run();

            // Cleanup
            timer.Stop();
            UnhookWindowsHookEx(hookId);
            Log("Exited.");
        }

        // --- Hook callback: INSTANT, just sets flags ---
        static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var kbd = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                if (kbd.vkCode == HOTKEY_VK)
                {
                    int msg = wParam.ToInt32();
                    if ((msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN) && !isRecording)
                    {
                        isRecording = true;
                        startFlag = true;
                    }
                    else if ((msg == WM_KEYUP || msg == WM_SYSKEYUP) && isRecording)
                    {
                        isRecording = false;
                        stopFlag = true;
                    }
                    return (IntPtr)1; // swallow key
                }
            }
            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }

        // --- Timer tick: does heavy work on UI thread ---
        static void OnTimerTick(object sender, EventArgs e)
        {
            if (startFlag)
            {
                startFlag = false;
                StartRecording();
            }
            if (stopFlag)
            {
                stopFlag = false;
                StopRecordingAndTranscribe();
            }
        }

        // --- Recording ---
        static void StartRecording()
        {
            Log("Recording START");
            try
            {
                var fmt = new WAVEFORMATEX();
                fmt.wFormatTag = 1;
                fmt.nChannels = CHANNELS;
                fmt.nSamplesPerSec = SAMPLE_RATE;
                fmt.wBitsPerSample = BITS;
                fmt.nBlockAlign = (ushort)(CHANNELS * BITS / 8);
                fmt.nAvgBytesPerSec = (uint)(SAMPLE_RATE * fmt.nBlockAlign);

                int r = waveInOpen(out hwi, -1, ref fmt, IntPtr.Zero, IntPtr.Zero, 0);
                if (r != 0) { Log("ERROR: waveInOpen=" + r); return; }

                bufPtr = Marshal.AllocHGlobal(bufSize);
                int hs = Marshal.SizeOf(typeof(WAVEHDR));
                hdrPtr = Marshal.AllocHGlobal(hs);
                var hdr = new WAVEHDR { lpData = bufPtr, dwBufferLength = (uint)bufSize };
                Marshal.StructureToPtr(hdr, hdrPtr, false);

                waveInPrepareHeader(hwi, hdrPtr, hs);
                waveInAddBuffer(hwi, hdrPtr, hs);
                waveInStart(hwi);

                ThreadPool.QueueUserWorkItem(_ => { try { Beep(800, 120); } catch { } });
            }
            catch (Exception ex) { Log("ERROR StartRecording: " + ex.Message); }
        }

        static void StopRecordingAndTranscribe()
        {
            Log("Recording STOP");
            try
            {
                if (hwi == IntPtr.Zero) { Log("ERROR: not recording"); return; }

                waveInStop(hwi);
                waveInReset(hwi);

                var hdr = (WAVEHDR)Marshal.PtrToStructure(hdrPtr, typeof(WAVEHDR));
                uint recorded = hdr.dwBytesRecorded;
                byte[] pcm = null;
                if (recorded >= 1600)
                {
                    pcm = new byte[recorded];
                    Marshal.Copy(bufPtr, pcm, 0, (int)recorded);
                }

                // Cleanup waveIn
                int hs = Marshal.SizeOf(typeof(WAVEHDR));
                waveInUnprepareHeader(hwi, hdrPtr, hs);
                waveInClose(hwi);
                hwi = IntPtr.Zero;
                Marshal.FreeHGlobal(hdrPtr); hdrPtr = IntPtr.Zero;
                Marshal.FreeHGlobal(bufPtr); bufPtr = IntPtr.Zero;

                if (pcm == null) { Log("Too short, skip"); return; }

                WriteWav(wavPath, pcm);

                ThreadPool.QueueUserWorkItem(_ => { try { Beep(400, 120); } catch { } });

                // Transcribe on background thread
                ThreadPool.QueueUserWorkItem(_ => Transcribe(wavPath));
            }
            catch (Exception ex) { Log("ERROR StopRecording: " + ex.Message); }
        }

        // --- WAV writer ---
        static void WriteWav(string path, byte[] pcm)
        {
            int byteRate = SAMPLE_RATE * CHANNELS * BITS / 8;
            int blockAlign = CHANNELS * BITS / 8;
            using (var fs = new FileStream(path, FileMode.Create))
            using (var bw = new BinaryWriter(fs))
            {
                bw.Write(new[] { 'R', 'I', 'F', 'F' }); bw.Write(36 + pcm.Length);
                bw.Write(new[] { 'W', 'A', 'V', 'E' }); bw.Write(new[] { 'f', 'm', 't', ' ' });
                bw.Write(16); bw.Write((short)1); bw.Write((short)CHANNELS); bw.Write(SAMPLE_RATE);
                bw.Write(byteRate); bw.Write((short)blockAlign); bw.Write((short)BITS);
                bw.Write(new[] { 'd', 'a', 't', 'a' }); bw.Write(pcm.Length); bw.Write(pcm);
            }
        }

        // --- Groq Whisper API ---
        static void Transcribe(string wavFile)
        {
            Log("Transcribing...");
            try
            {
                byte[] fileBytes = File.ReadAllBytes(wavFile);
                string boundary = Guid.NewGuid().ToString();

                // Build multipart body as bytes (binary safe)
                var body = new List<byte>();
                string header =
                    "--" + boundary + "\r\n" +
                    "Content-Disposition: form-data; name=\"file\"; filename=\"audio.wav\"\r\n" +
                    "Content-Type: audio/wav\r\n\r\n";
                body.AddRange(Encoding.UTF8.GetBytes(header));
                body.AddRange(fileBytes);
                string middle =
                    "\r\n--" + boundary + "\r\n" +
                    "Content-Disposition: form-data; name=\"model\"\r\n\r\n" +
                    WHISPER_MODEL +
                    "\r\n--" + boundary + "--\r\n";
                body.AddRange(Encoding.UTF8.GetBytes(middle));

                byte[] bodyBytes = body.ToArray();

                var req = (HttpWebRequest)WebRequest.Create("https://api.groq.com/openai/v1/audio/transcriptions");
                req.Method = "POST";
                req.ContentType = "multipart/form-data; boundary=" + boundary;
                req.Headers["Authorization"] = "Bearer " + GROQ_API_KEY;
                req.ContentLength = bodyBytes.Length;
                req.Timeout = 30000;

                using (var s = req.GetRequestStream())
                    s.Write(bodyBytes, 0, bodyBytes.Length);

                string responseText;
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                    responseText = sr.ReadToEnd();

                // Parse JSON manually (no external deps)
                string text = ExtractJsonText(responseText);
                if (string.IsNullOrEmpty(text))
                {
                    Log("No text in response");
                    return;
                }
                Log("Transcription: '" + text + "'");

                // Paste via clipboard + Ctrl+V (Unicode safe)
                PasteText(text);
            }
            catch (WebException wex)
            {
                string body = "";
                try
                {
                    if (wex.Response != null)
                        using (var sr = new StreamReader(wex.Response.GetResponseStream()))
                            body = sr.ReadToEnd();
                }
                catch { }
                Log("ERROR API: " + wex.Message + " | " + body);
            }
            catch (Exception ex)
            {
                Log("ERROR Transcribe: " + ex.Message);
            }
        }

        static string ExtractJsonText(string json)
        {
            // Find "text":"..." in JSON response
            int idx = json.IndexOf("\"text\"");
            if (idx < 0) return null;
            idx = json.IndexOf(":", idx);
            if (idx < 0) return null;
            idx = json.IndexOf("\"", idx + 1);
            if (idx < 0) return null;
            int start = idx + 1;
            int end = start;
            while (end < json.Length)
            {
                if (json[end] == '\\') { end += 2; continue; }
                if (json[end] == '"') break;
                end++;
            }
            string raw = json.Substring(start, end - start);
            // Unescape basic JSON escapes
            raw = raw.Replace("\\n", "\n").Replace("\\r", "\r").Replace("\\t", "\t")
                     .Replace("\\\"", "\"").Replace("\\\\", "\\");
            return raw.Trim();
        }

        // --- Paste text via clipboard + SendInput(Ctrl+V) ---
        static void PasteText(string text)
        {
            // Must run clipboard ops on STA thread
            var t = new Thread(() =>
            {
                try
                {
                    string oldClip = null;
                    try { oldClip = Clipboard.GetText(); } catch { }

                    Clipboard.SetText(text);
                    Thread.Sleep(100);

                    // SendInput Ctrl+V
                    var inputs = new INPUT[4];
                    int sz = Marshal.SizeOf(typeof(INPUT));
                    inputs[0].type = 1; inputs[0].u.ki.wVk = 0x11;
                    inputs[1].type = 1; inputs[1].u.ki.wVk = 0x56;
                    inputs[2].type = 1; inputs[2].u.ki.wVk = 0x56; inputs[2].u.ki.dwFlags = 2;
                    inputs[3].type = 1; inputs[3].u.ki.wVk = 0x11; inputs[3].u.ki.dwFlags = 2;
                    SendInput(4, inputs, sz);
                    Log("Pasted: '" + text + "'");

                    Thread.Sleep(200);

                    // Restore clipboard
                    if (!string.IsNullOrEmpty(oldClip))
                        Clipboard.SetText(oldClip);
                    else
                        Clipboard.Clear();
                }
                catch (Exception ex) { Log("ERROR Paste: " + ex.Message); }
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        // --- Logging ---
        static void Log(string msg)
        {
            string line = string.Format("[{0:yyyy-MM-dd HH:mm:ss.fff}] {1}", DateTime.Now, msg);
            Console.WriteLine(line);
            try { File.AppendAllText(logPath, line + Environment.NewLine); } catch { }
        }
    }
}
