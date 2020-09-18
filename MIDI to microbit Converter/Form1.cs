using System;
using System.IO;
using com.bluejay113.tonelibrary;
using System.Threading;
using System.Drawing;
using System.Windows.Forms;
using Melanchall.DryWetMidi.Core;
using Melanchall.DryWetMidi.Interaction;
using System.Linq;
using System.Diagnostics;

namespace MIDI_to_microbit_Converter
{
    public partial class Form1 : Form
    {
        internal toneDictionary noteDict { get; set; } = new toneDictionary();
        public int pageNum = -1;
        public Random rand = new Random();
        public string midiFilePath = "";
        public int COMchannel = 0;
        public string activationCode = "";
        public string result = "";

        public SplitContainer TopMiddlecontainer = new SplitContainer();
        public SplitContainer MiddleBottomcontainer = new SplitContainer();
        public Label title = new Label();
        public PictureBox logo = new PictureBox();
        public Button next = new Button();
        public Button back = new Button();
        public Button cancel = new Button();

        public Label page0Instructions = new Label();

        public Label page1Instructions = new Label();
        public Label midiFileNameLabel = new Label();
        public Button importMIDI = new Button();
        public OpenFileDialog importMIDIDialog = new OpenFileDialog();
        public CheckBox multipleMicroBits = new CheckBox();
        public GroupBox COMSettingsTable = new GroupBox();
        public Label COMwarning = new Label();
        public Label COMchannelLabel = new Label();
        public NumericUpDown COMchannelBox = new NumericUpDown();
        public Label activationCodeLabel = new Label();
        public TextBox activationCodeTextBox = new TextBox();

        public Label progressText = new Label();
        public ProgressBar conversionProgress = new ProgressBar();

        public LinkLabel page3Instructions = new LinkLabel();
        public WebBrowser convertedResult = new WebBrowser();

        public Form1()
        {
            InitializeComponent();
            noteDict.Init();
            // Reference Form Controls
            TopMiddlecontainer = new SplitContainer() { BackColor = Color.White, BorderStyle = BorderStyle.None, Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, IsSplitterFixed = true, FixedPanel = FixedPanel.Panel1, Panel1MinSize = 100, SplitterWidth = 1 };
            MiddleBottomcontainer = new SplitContainer() { BackColor = SystemColors.Control, BorderStyle = BorderStyle.None, Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, IsSplitterFixed = true, FixedPanel = FixedPanel.Panel2, Panel2MinSize = 30 };
            title = new Label() { TextAlign = ContentAlignment.MiddleLeft, Text = "MIDI to micro:bit Converter", Margin = new Padding(20,20,0,0), Font = new Font(new FontFamily("Corbel"), 14, FontStyle.Regular), AutoSize = true, Location = new Point(10, TopMiddlecontainer.Panel1.Height / 3) };
            logo = new PictureBox() { Image = Properties.Resources.logo, SizeMode = PictureBoxSizeMode.StretchImage, Dock = DockStyle.Right, Height = TopMiddlecontainer.Panel1.Height, Width = TopMiddlecontainer.Panel1.Height, Margin = new Padding(0,0,0,0) };
            next = new Button() { Text = " Next > ", AutoSize = false, Location = new Point(420, MiddleBottomcontainer.Panel2.Height / 4), Height = MiddleBottomcontainer.Panel2.Height / 2 };
            back = new Button() { Text = " < Back ", AutoSize = false, Location = new Point(340, MiddleBottomcontainer.Panel2.Height / 4), Height = MiddleBottomcontainer.Panel2.Height / 2 };
            cancel = new Button() { Text = " Cancel ", AutoSize = false, Location = new Point(500, MiddleBottomcontainer.Panel2.Height / 4), Height = MiddleBottomcontainer.Panel2.Height / 2 };
            
            // Start Page
            page0Instructions = new Label() { Height = MiddleBottomcontainer.Panel1.Height - 40, Width = MiddleBottomcontainer.Panel1.Width - 40, Dock = DockStyle.Fill, Padding = new Padding(20, 0, 20, 0), TextAlign = ContentAlignment.MiddleLeft, Text = "Welcome to the MIDI to micro:bit Converter!\n\nThis tool converts MIDI files to Python code which the BBC® micro:bit© can use.\n\nAt this moment, this tool can only extract notes from a single instrument. If you import a MIDI with more than 1 instrument, this tool will end up placing every note from each instrument into 1 instrument (It doesn't sound good so don't do it).\n\nIf you need to use multiple instruments (to make a fuller sound), split every instrument into its own MIDI file, process it through this tool, and write the output code to the micro:bit© IDE.\n\nThe micro:bit© uses frequency tones (in Hertz) to make sound. The micro:bit© have no digital audio support (MP3, FLAC, WAV, etc.), only single instrument MIDIs (per micro:bit).\n\nNote: micro:bit© pushes mono (single channel) audio through pin 0 (P0). If you use a MIDI which has a very long song length, it will take a while for the program to compile.\n\nPress Next to continue." };

            // Page 1
            page1Instructions = new Label() { Dock = DockStyle.Top, Text = "Choose the MIDI file below you want to convert.", Padding = new Padding(20,20,0,0), AutoSize = true };
            importMIDI = new Button() { Text = "Import MIDI", AutoSize = true, Location = new Point(20, 40) };
            midiFileNameLabel = new Label() { AutoSize = true, Location = new Point(100, 44) };
            multipleMicroBits = new CheckBox() { Appearance = Appearance.Normal, Text = "Sync multiple micro:bit©s", Location = new Point(20, 70), AutoSize = true };
            COMSettingsTable = new GroupBox() { Text = "micro:bit Communication Settings", AutoSize = true, Location = new Point(20, 100) };
            COMwarning = new Label() { Text = "Note: You must keep the COM Channel and Activation Code the same for each micro:bit©.", Location = new Point(5, 20), Font = new Font(DefaultFont.FontFamily, DefaultFont.Size, FontStyle.Bold), ForeColor = Color.Red, AutoSize = true };
            COMchannelLabel = new Label() { Text = "COM Channel:", AutoSize = true, Location = new Point(20, 48) };
            COMchannelBox = new NumericUpDown() { Increment = 1, Minimum = 0, Maximum = 99, Width = 60, Location = new Point(100, 45) };
            activationCodeLabel = new Label() { Text = "Activation Code (for sync):", AutoSize = true, Location = new Point(20, 70) };
            activationCodeTextBox = new TextBox() { Width = 200, Multiline = false, MaxLength = 30, Location = new Point(155, 68) };

            // Page 2
            progressText = new Label() { Location = new Point(MiddleBottomcontainer.Panel1.Width / 3, MiddleBottomcontainer.Panel1.Height / 3) };
            conversionProgress = new ProgressBar() { Height = 50, Width = MiddleBottomcontainer.Panel1.Width, Location = new Point(0, MiddleBottomcontainer.Panel1.Height / 3 - 50), Value = 100, MarqueeAnimationSpeed = 200 };

            // Page 3
            page3Instructions = new LinkLabel() { Text = "The conversion process has completed! Please go to https://makecode.microbit.org and\npaste the code which has been given below. Thank you for using this tool!", LinkArea = new LinkArea(51, 29), LinkBehavior = LinkBehavior.AlwaysUnderline, AutoSize = true, Padding = new Padding(20) };
            page3Instructions.Links[0].Description = "https://makecode.microbit.org";
            convertedResult = new WebBrowser() { AllowNavigation = false, AllowWebBrowserDrop = false, Dock = DockStyle.Bottom, Height = 200, ScriptErrorsSuppressed = true, ScrollBarsEnabled = true };

            // Events
            TopMiddlecontainer.Panel1.Paint += Panel1_Paint;
            MiddleBottomcontainer.Panel2.Paint += Panel2_Paint;
            cancel.Click += Cancel_Click;
            next.Click += Next_Click;
            back.Click += Back_Click;
            importMIDI.Click += ImportMIDI_Click;
            multipleMicroBits.CheckedChanged += MultipleMicroBits_CheckedChanged;
            page3Instructions.LinkClicked += Page3Instructions_LinkClicked;

            page("+");
        }

        private void Page3Instructions_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(e.Link.Description);
        }

        private void MultipleMicroBits_CheckedChanged(object sender, EventArgs e)
        {
            if (multipleMicroBits.CheckState == CheckState.Checked)
            {
                MiddleBottomcontainer.Panel1.Controls.Add(COMSettingsTable);
                COMSettingsTable.Controls.Add(COMwarning);
                COMSettingsTable.Controls.Add(COMchannelLabel);
                COMSettingsTable.Controls.Add(COMchannelBox);
                COMSettingsTable.Controls.Add(activationCodeLabel);
                COMSettingsTable.Controls.Add(activationCodeTextBox);
            }
            else
            {
                MiddleBottomcontainer.Panel1.Controls.Remove(COMSettingsTable);
                COMSettingsTable.Controls.Remove(COMwarning);
                COMSettingsTable.Controls.Remove(COMchannelLabel);
                COMSettingsTable.Controls.Remove(COMchannelBox);
                COMSettingsTable.Controls.Remove(activationCodeLabel);
                COMSettingsTable.Controls.Remove(activationCodeTextBox);
            }
        }

        private void ImportMIDI_Click(object sender, EventArgs e)
        {
            importMIDIDialog.Filter = "MIDI files (*.mid, *.midi)|*.mid;*.midi";
            importMIDIDialog.FilterIndex = 0;
            if (importMIDIDialog.ShowDialog() == DialogResult.OK)
            {
                midiFilePath = importMIDIDialog.FileName;
                midiFileNameLabel.Text = importMIDIDialog.SafeFileName;
                next.Enabled = true;
            }
        }

        private void Back_Click(object sender, EventArgs e)
        {
            page("-");
            if (pageNum == 0)
            {
                back.Enabled = false;
            }

            if (pageNum < 0)
            {
                pageNum = 0;
            }
        }

        private void Next_Click(object sender, EventArgs e)
        {
            if (next.Text != "Finish")
            {
                page("+");
                if (pageNum == 1)
                {
                    back.Enabled = true;
                }
            }
            else
            {
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\AppData\Local\Temp\result.html");
                Application.Exit();
            }
        }

        private void Cancel_Click(object sender, System.EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to quit?", "MIDI to micro:bit Converter", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes) { try { Application.Exit(); } catch (Exception) { Thread.Sleep(10); Application.Exit(); } }
        }

        private void Panel1_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, e.ClipRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void Panel2_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, e.ClipRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        public void page(string addminus)
        {
            var elements = MiddleBottomcontainer.Panel1.Controls.Count;
            if (elements > 0)
            {
                MiddleBottomcontainer.Panel1.Controls.Clear();
            }

            if (addminus == "+") { pageNum++; } else if (addminus == "-") { pageNum--; }

            if (pageNum == 0)
            {
                next.Enabled = true;
                Controls.Add(TopMiddlecontainer);
                TopMiddlecontainer.Panel2.Controls.Add(MiddleBottomcontainer);
                MiddleBottomcontainer.Panel2.BackColor = Color.White;
                TopMiddlecontainer.Panel1.Controls.Add(title);
                TopMiddlecontainer.Panel1.Controls.Add(logo);
                MiddleBottomcontainer.Panel1.Controls.Add(page0Instructions);
                MiddleBottomcontainer.Panel2.Controls.Add(next);
                MiddleBottomcontainer.Panel2.Controls.Add(back);
                back.Enabled = false;
                MiddleBottomcontainer.Panel2.Controls.Add(cancel);
            }
            else if (pageNum == 1)
            {
                if (midiFilePath != "")
                {
                    next.Enabled = true;
                }
                else
                {
                    next.Enabled = false;
                    COMchannelBox.Value = rand.Next(0, 99);
                }
                if (multipleMicroBits.Checked == true)
                {
                    MiddleBottomcontainer.Panel1.Controls.Add(COMSettingsTable);
                    COMSettingsTable.Controls.Add(COMwarning);
                    COMSettingsTable.Controls.Add(COMchannelLabel);
                    COMSettingsTable.Controls.Add(COMchannelBox);
                    COMSettingsTable.Controls.Add(activationCodeLabel);
                    COMSettingsTable.Controls.Add(activationCodeTextBox);
                }
                MiddleBottomcontainer.Panel1.Controls.Add(page1Instructions);
                MiddleBottomcontainer.Panel1.Controls.Add(importMIDI);
                MiddleBottomcontainer.Panel1.Controls.Add(midiFileNameLabel);
                MiddleBottomcontainer.Panel1.Controls.Add(multipleMicroBits);
            }
            else if (pageNum == 2)
            {
                MiddleBottomcontainer.Panel1.Controls.Add(progressText);
                MiddleBottomcontainer.Panel1.Controls.Add(conversionProgress);
                if (multipleMicroBits.Checked == true)
                {
                    if (activationCodeTextBox.Text == "" || activationCodeTextBox.Text.Contains(" ") == true)
                    {
                        MessageBox.Show("Invalid Activation Code.\n\nMake sure there are no spaces and that there is a value in the code box.", "MIDI to micro:bit© Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        page("-");
                    }
                }
                COMchannel = (int)COMchannelBox.Value;
                activationCode = activationCodeTextBox.Text;
                if (activationCodeTextBox.Text != "")
                {
                    result = "brightness = 0<br>height = 0<br>width = 0<br>bud = True<br>playSong = False<br>radio.set_group(" + COMchannel + ")<br><br>basic.clear_screen()<br>def on_received_string(receivedString):<br>&nbsp&nbsp&nbsp&nbspif (receivedString == '" + activationCode + @"'):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (playSong == False):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspplaySong = True<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbasic.clear_screen()<br>radio.on_received_string(on_received_string)<br><br>while True:<br>&nbsp&nbsp&nbsp&nbspif (input.button_is_pressed(Button.A) == True):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspradio.send_string('" + activationCode + "')<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (playSong == False):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspplaySong = True<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbasic.clear_screen()<br>&nbsp&nbsp&nbsp&nbspelif (input.button_is_pressed(Button.B) == True):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (playSong == True):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspplaySong = False<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbasic.clear_screen()<br>&nbsp&nbsp&nbsp&nbspif (playSong == False):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (brightness > 255):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbud = False<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspelif (brightness == 0):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbud = True<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (width == 5):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspheight = height + 1<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspwidth = 0<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (height == 5):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspheight = 0<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (bud == True):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbrightness = brightness + 1<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspelif (bud == False):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbrightness = brightness - 1<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbasic.pause(0)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot_brightness(width, height, brightness)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspwidth = width + 1<br>&nbsp&nbsp&nbsp&nbspelif (playSong == True):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(1, 0)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(1, 1)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(1, 2)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(1, 3)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(1, 4)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(2, 1)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(2, 2)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(2, 3)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(3, 2)<br>";
                }
                else
                {
                    result = "brightness = 0<br>height = 0<br>width = 0<br>bud = True<br>playSong = False<br><br>basic.clear_screen()<br>while True:<br>&nbsp&nbsp&nbsp&nbspif (input.button_is_pressed(Button.A) == True):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (playSong == False):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspplaySong = True<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbasic.clear_screen()<br>&nbsp&nbsp&nbsp&nbspelif (input.button_is_pressed(Button.B) == True):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (playSong == True):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspplaySong = False<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbasic.clear_screen()<br>&nbsp&nbsp&nbsp&nbspif (playSong == False):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (brightness > 255):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbud = False<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspelif (brightness == 0):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbud = True<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (width == 5):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspheight = height + 1<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspwidth = 0<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (height == 5):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspheight = 0<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspif (bud == True):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbrightness = brightness + 1<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspelif (bud == False):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbrightness = brightness - 1<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspbasic.pause(0)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot_brightness(width, height, brightness)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspwidth = width + 1<br>&nbsp&nbsp&nbsp&nbspelif (playSong == True):<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(1, 0)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(1, 1)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(1, 2)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(1, 3)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(1, 4)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(2, 1)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(2, 2)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(2, 3)<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspled.plot(3, 2)<br>";
                }
                next.Enabled = false;
                back.Enabled = false;
                cancel.Enabled = false;
                var midi = MidiFile.Read(midiFilePath);
                var tempoMap = midi.GetTempoMap();
                MetricTimeSpan metricTime = midi.GetTimedEvents().Skip(9).First().TimeAs<MetricTimeSpan>(tempoMap);
                BarBeatTicksTimeSpan tempoInt = TimeConverter.ConvertTo<BarBeatTicksTimeSpan>(metricTime, tempoMap);
                result = string.Join("", result + "<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspmusic.set_volume(80)");
                var notesManager = midi.GetTrackChunks().First().ManageNotes();
                var numNotes = notesManager.Notes.ToList().Count;
                var currentNote = 0;
                foreach (var note in midi.GetNotes())
                {
                    var noteName = note.GetMusicTheoryNote().ToString();
                    var noteTime = note.LengthAs<MetricTimeSpan>(tempoMap).Milliseconds;
                    var notePitch = "";
                    foreach (var item in noteDict.Note.Where(i => i.Key == noteName))
                    {
                        notePitch = string.Join("", item.Value);
                    }
                    result = string.Join("", result + "<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspmusic.play_tone(" + notePitch + ", " + noteTime + ")");
                    currentNote++;
                    progressText.Text = "Processing note " + currentNote + " of " + numNotes + "...";
                }
                page("+");
                result = string.Join("", result + "<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspplaySong = false");
            }
            else if (pageNum == 3)
            {
                var document = "<html><head><style>p{font-family: Verdana, Arial, Tahoma, Serif;font-size: 11px;}</style></head><body><p>" + result + "</p></body></html>";
                File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\AppData\Local\Temp\result.html", document);
                convertedResult.DocumentText = document;
                next.Text = "Finish";
                next.Enabled = true;
                MiddleBottomcontainer.Panel1.Controls.Add(page3Instructions);
                MiddleBottomcontainer.Panel1.Controls.Add(convertedResult);
            }
        }
    }
}
