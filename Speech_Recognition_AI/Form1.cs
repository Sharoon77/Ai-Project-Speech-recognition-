using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Speech;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Threading;
using System.Diagnostics;


namespace Speech_Recognition_AI
{
    public partial class Form1 : Form
    {
        SpeechSynthesizer ss = new SpeechSynthesizer();
        PromptBuilder pb = new PromptBuilder();
        SpeechRecognitionEngine sre = new SpeechRecognitionEngine();

        static string recent = "";
        
        public Form1()
        {

            InitializeComponent();
        }

        public void button1_Click(object sender, EventArgs e)
        {
            btnStart.Enabled = false;
            btnStop.Enabled = true;
            Choices clist = new Choices();
            clist.Add(new string[] { "hello john", "start please","what do you do","who made you","what is date today", "how are you","who are you","jarvis what you can do","what is day today", "open chrome", "open facebook", "open youtube", "open word", "open excel", "open notepad", "open paint", "open calculator", "open camera", "open pc", "open c", "open recent file", "close chrome", "close facebook", "close youtube", "close notepad", "close word", "close excel", "close paint", "close calculator", "close please", "close camera", "close pc", "close c" });
            Grammar gr = new Grammar(new GrammarBuilder(clist));

            try
            {
                sre.RequestRecognizerUpdate();
                sre.LoadGrammar(gr);
                sre.SpeechRecognized += sre_SpeechRecognized;
                sre.SetInputToDefaultAudioDevice();
                sre.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch { }
        }

        private void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch(e.Result.Text.ToString())
            {
                case "hello john":
                    ss.SpeakAsync("Hello how can i help you ?");
                    break;
                case "what do you do":
                    ss.SpeakAsync("scooby cooby do");
                    break;
                case "how are you":
                    ss.SpeakAsync("I am fine");
                    break;
                case "who made you":
                    ss.SpeakAsync("Sharoon is my creator");
                    break;
                case "who are you":
                    ss.SpeakAsync("I am your personal Asistant");
                    break;
                case "jarvis what you can do":
                    ss.SpeakAsync(" Sir I responce to your command");
                    break;
                case "what is day today":
                    string a = System.DateTime.Now.DayOfWeek.ToString();
                    ss.SpeakAsync(a);
                    break;
                case "what is date today":
                    DateTime dateTime = DateTime.UtcNow.Date;
                    ss.SpeakAsync(dateTime.ToString("dd/MM/yyyy"));
                    break;
                case "open chrome":
                    var prs = new ProcessStartInfo("chrome.exe");
                    prs.Arguments = "http://google.com";
                    Process.Start(prs);
                    break;
                case "open facebook":
                    var prs1 = new ProcessStartInfo("chrome.exe");
                    prs1.Arguments = "http://facebook.com";
                    Process.Start(prs1);
                    break;
                case "open youtube":
                    var prs2 = new ProcessStartInfo("chrome.exe");
                    prs2.Arguments = "http://youtube.com";
                    Process.Start(prs2);
                    break;
                case "open word":
                    Process.Start("winword");
                    recent = "winword";
                    break;
                case "open excel":
                    Process.Start("excel");
                    recent = "excel";
                    break;
                case "open paint":
                    Process.Start("mspaint");
                    recent = "mspaint";
                    break;
                case "open notepad":
                    Process.Start("notepad");
                    recent = "notpad";
                    break;
                case "open calculator":
                    Process.Start("calc");
                    recent = "calc";
                    break;
                case "open camera":
                    Process.Start("microsoft.windows.camera:");
                    recent = "microsoft.windows.camera:";
                    break;
                case "open pc":
                    Process.Start("::{20d04fe0-3aea-1069-a2d8-08002b30309d}");
                    recent = "::{20d04fe0-3aea-1069-a2d8-08002b30309d}";
                    break;
                case "open c":
                    Process.Start(@"c:\Users");
                    recent = "@\"c:\\Users\"";
                    break;
                case "open recent file":
                    if (recent == "")
                    {
                        ss.SpeakAsync("Sorry there is no recent files there");
                        break;
                    }
                    else
                    {
                        Process.Start(recent);
                        break;
                    }
                case "close notepad":
                    foreach (var process in Process.GetProcessesByName("notepad"))
                    {
                        process.Kill();
                    }
                    break;
                case "close word":
                    foreach (var process in Process.GetProcessesByName("winword"))
                    {
                        process.Kill();
                    }
                    break;
                case "close excel":
                    foreach (var process in Process.GetProcessesByName("excel"))
                    {
                        process.Kill();
                    }
                    break;
                case "close camera":
                    foreach (var process in Process.GetProcessesByName("microsoft.windows.camera:"))
                    {
                        process.Kill();
                    }
                    break;
                
                case "close paint":
                    foreach (var process in Process.GetProcessesByName("mspaint"))
                    {
                        process.Kill();
                    }
                    break;
                case "close calculator":
                    foreach (var process in Process.GetProcessesByName("calc"))
                    {
                        process.Kill();
                    }
                    break;
                case "close pc":
                    foreach (var process in Process.GetProcessesByName("::{20d04fe0-3aea-1069-a2d8-08002b30309d}"))
                    {
                        process.Kill();
                    }
                    break;
                case "close c":
                    foreach (var process in Process.GetProcessesByName("*.exe"))
                    {
                        process.Kill();
                    }
                    break;
                case "close please":
                    this.Close();
                    break;
                default:
                    ss.SpeakAsync("Sorry I didnt Get it");
                    break;
                    
            }
          txtContents.Text += e.Result.Text.ToString() + Environment.NewLine; 
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            sre.RecognizeAsyncStop();
            btnStart.Enabled = true;
            btnStop.Enabled = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
        
    }
}
