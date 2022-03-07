


// -------------------------------------------------------------------------------------------

// Copyright (c) 2020 Federico Turco

// Permission is hereby granted, free of charge, to any person
// obtaining a copy of this software and associated documentation
// files (the "Software"), to deal in the Software without
// restriction, including without limitation the rights to use,
// copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the
// Software is furnished to do so, subject to the following
// conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
// OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
// HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
// WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
// OTHER DEALINGS IN THE SOFTWARE.

// -------------------------------------------------------------------------------------------

// NB: I file in pdf accessibili dal menu info sono di proprietà dei rispettivi autori

// -------------------------------------------------------------------------------------------


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
using System.Reflection;

// Sockets
using System.Net.Sockets;

// Ping
using System.Net.NetworkInformation;

// Componenti quali MessageBox
//using System.Windows.Forms;

// Threading per server ModBus TCP
using System.Threading;

using System.Collections;

// Porta seriale
using System.IO.Ports;

// Comandi apri/chiudi console
using System.Runtime.InteropServices;

using System.Runtime.Serialization;

// Libreria JSON
using System.IO;

using ModBusMaster_Chicco;

//using System.Windows.Media;
using System.Collections.ObjectModel;

using Microsoft.Win32;
using System.Diagnostics;

// Libreria JSON
using System.Web.Script.Serialization;

// Libreria lingue
using LanguageLib; // Libreria custom per caricare etichette in lingue differenti

namespace ModBus_Server
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // NOTE Icona
        // Icona -> va messa nel Form1.Designer dopo averla caricata in Proprieta' -> Risorse
        // this.Icon = MQTT_Almaviva.Properties.Resources.IconaAlgorab;

        //-----------------------------------------------------
        //-----------------Variabili globali-------------------
        //-----------------------------------------------------

        String version = "beta";
        String title = "ModBus C#";

        String defaultPathToConfiguration = "Generico";
        String pathToConfiguration;

        // Colori default caselle
        public SolidColorBrush colorDefaultReadCell = Brushes.LightSkyBlue;
        public SolidColorBrush colorDefaultWriteCell = Brushes.LightGreen;
        public SolidColorBrush colorErrorCell = Brushes.Orange;
        public SolidColorBrush colorMenuStrip;

        public ObservableCollection<ModBus_Item> list_coilsTable = new ObservableCollection<ModBus_Item>();
        public ObservableCollection<ModBus_Item> list_inputsTable = new ObservableCollection<ModBus_Item>();
        public ObservableCollection<ModBus_Item> list_inputRegistersTable = new ObservableCollection<ModBus_Item>();
        public ObservableCollection<ModBus_Item> list_holdingRegistersTable = new ObservableCollection<ModBus_Item>();

        System.Windows.Forms.ColorDialog colorDialogBox = new System.Windows.Forms.ColorDialog();

        // Elementi per visualizzare/nascondere la finestra della console
        bool statoConsole = true;

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        // Disable Console Exit Button
        [DllImport("user32.dll")]
        static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        [DllImport("user32.dll")]
        static extern IntPtr DeleteMenu(IntPtr hMenu, uint uPosition, uint uFlags);

        const uint SC_CLOSE = 0xF060;
        const uint MF_BYCOMMAND = (uint)0x00000000L;

        public ModBus_Chicco ModBus;

        public SerialPort serialPort = new SerialPort();
        public string stringSerialPort = "COMxx";

        //Threads per programmi in background
        Thread loopTcp;
        Thread loopSerial;

        SaveFileDialog saveFileDialogBox;
        OpenFileDialog openFileDialogBox;

        Thread threadDequeue;

        public string language = "IT";

        Language lang;

        public bool logWindowIsOpen = false;
        public bool dequeueExit = false;

        public int LogLimitRichTextBox = 2000;

        public bool ShowAlertEditMode = false;

        bool scrolled_log = false;
        int count_log = 0;

        bool scrolled_status = false;
        int count_status = 0;

        public MainWindow()
        {
            InitializeComponent();

            lang = new Language(this);

            // Menu lingua
            languageToolStripMenu.Items.Clear();

            foreach (string lang in Directory.GetFiles(Directory.GetCurrentDirectory() + "//Lang"))
            {
                var tmp = new MenuItem();

                tmp.Header = System.IO.Path.GetFileNameWithoutExtension(lang);
                tmp.IsCheckable = true;
                tmp.Click += MenuItemLanguage_Click;

                languageToolStripMenu.Items.Add(tmp);
            }

            lang.loadLanguageTemplate(language);

            pathToConfiguration = defaultPathToConfiguration;

            pictureBoxSerial.Background = Brushes.LightGray;
            pictureBoxTcp.Background = Brushes.LightGray;

            dataGridViewCoils.ItemsSource = list_coilsTable;
            dataGridViewInput.ItemsSource = list_inputsTable;
            dataGridViewInputRegister.ItemsSource = list_inputRegistersTable;
            dataGridViewHolding.ItemsSource = list_holdingRegistersTable;

            // Aspetti grafici di default
            comboBoxSerialSpeed.SelectedIndex = 7;
            comboBoxSerialParity.SelectedIndex = 0;
            comboBoxSerialStop.SelectedIndex = 0;

            textBoxTcpClientIpAddress.Text = "192.168.1.100";
            textBoxTcpClientPort.Text = "502";

            comboBoxCoilsRegisters.SelectedIndex = 0;
            comboBoxCoilsOffset.SelectedIndex = 0;
            textBoxCoilsOffset.Text = "0";

            comboBoxInputsRegisters.SelectedIndex = 0;
            comboBoxInputsOffset.SelectedIndex = 0;
            textBoxInputsOffset.Text = "0";

            comboBoxInputsRegRegisters.SelectedIndex = 0;
            comboBoxInputsRegValues.SelectedIndex = 0;
            comboBoxInputsRegOffset.SelectedIndex = 0;
            textBoxInputsRegOffset.Text = "0";

            comboBoxHoldingsRegisters.SelectedIndex = 0;
            comboBoxHoldingsValues.SelectedIndex = 0;
            comboBoxHoldingsOffset.SelectedIndex = 0;
            textBoxHoldingsOffset.Text = "0";

            pictureBoxRunningAs.Background = Brushes.LightGray;
            pictureBoxIsSending.Background = Brushes.LightGray;
            pictureBoxIsResponding.Background = Brushes.LightGray;

            richTextBoxStatus.AppendText("\n");

            radioButtonModeSerial.IsChecked = true;

            checkBoxAddLinesToEnd.Visibility = Visibility.Hidden;

            updateStartPauseStop();

            // Disabilita il pulsante di chiusura della console
            disableConsoleExitButton();
        }

        
        public void MenuItemLanguage_Click(object sender, EventArgs e)
        {
            var currMenuItem = (MenuItem)sender;

            language = currMenuItem.Header.ToString();

            // Passo fuori le lingue disponibili nel menu
            foreach (MenuItem tmp in languageToolStripMenu.Items)
            {
                tmp.IsChecked = tmp == currMenuItem;
            }

            // Carico template selezionato
            lang.loadLanguageTemplate(currMenuItem.Header.ToString());
        }

        //------------------------------------------------------------------------
        //----------------Funzione chiamata alla chiusura del form----------------
        //------------------------------------------------------------------------

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            salva_configurazione(false);

            try
            {
                if (ModBus != null)
                {
                    ModBus.stopTcpLoop();
                }
            }
            catch { }

            try
            {
                if (loopTcp != null)
                {
                    if (loopTcp.IsAlive)
                    {
                        loopTcp.Abort();
                    }
                }
            }
            catch { }

            try
            {
                if (loopSerial != null)
                {
                    if (loopSerial.IsAlive)
                    {
                        loopSerial.Abort();
                    }
                }
            }
            catch { }
        }

        //-------------------------------------------------------------------------
        //----------------Funzione chiamata al caricamento del form----------------
        //-------------------------------------------------------------------------

        private void Form1_Load(object sender, RoutedEventArgs e)
        {
            threadDequeue = new Thread(new ThreadStart(LogDequeue));
            threadDequeue.IsBackground = true;
            threadDequeue.Start();

            richTextBoxPackets.Document.Blocks.Clear();
            richTextBoxPackets.AppendText("\n");
            richTextBoxPackets.Document.PageWidth = 5000;

            var version_ = Assembly.GetEntryAssembly().GetName().Version;

            this.Title = "ModBus Server " + version_.ToString().Split('.')[0] + "." + version_.ToString().Split('.')[1] + " " + version;

            Console.WriteLine(this.Title + "\n");

            try
            {
                //Aggiornamento lista porte seriale
                string[] SerialPortList = System.IO.Ports.SerialPort.GetPortNames();

                //comboBoxSerialPort.Items.Add("Seleziona porta seriale ...");
                foreach (String port_ in SerialPortList)
                {
                    comboBoxSerialPort.Items.Add(port_);
                }
                comboBoxSerialPort.SelectedIndex = 0;
            }
            catch
            {
                Console.WriteLine("Nessuna porta seriale trovata");
            }

            carica_configurazione();
        }

        private void radioButtonModeSerial_CheckedChanged(object sender, RoutedEventArgs e)
        {

            //Serial ON

            //radioButtonModeASCII.IsEnabled = radioButtonModeSerial.IsChecked;

            comboBoxSerialParity.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            comboBoxSerialPort.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            comboBoxSerialSpeed.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            comboBoxSerialStop.IsEnabled = (bool)radioButtonModeSerial.IsChecked;

            buttonUpdateSerialList.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            buttonSerialActive.IsEnabled = (bool)radioButtonModeSerial.IsChecked;


            //Tcp OFF
            //radioButtonTcpSlave.IsEnabled = !radioButtonModeSerial.IsChecked;

            buttonTcpActive.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;

            if ((bool)radioButtonModeSerial.IsChecked)
            {
                textBoxTcpClientIpAddress.IsEnabled = false;
                textBoxTcpClientPort.IsEnabled = false;
            }
            else
            {

                textBoxTcpClientIpAddress.IsEnabled = true;
                textBoxTcpClientPort.IsEnabled = true;

            }
        }



        private void buttonSerialActive_Click(object sender, RoutedEventArgs e)
        {
            if (pictureBoxSerial.Background == Brushes.LightGray)
            {
                //Attivazione comunicazione seriale
                pictureBoxSerial.Background = Brushes.Lime;
                pictureBoxRunningAs.Background = Brushes.Lime;


                buttonSerialActive.Content = "Disconnect";

                try
                {
                    //---------------------------------------------------------------------------------
                    //----------------------Apertura comunicazione seriale-----------------------------
                    //---------------------------------------------------------------------------------

                    // Create a new SerialPort object with default settings.
                    serialPort = new SerialPort();

                    serialPort.PortName = comboBoxSerialPort.SelectedItem.ToString();

                    serialPort.BaudRate = int.Parse(comboBoxSerialSpeed.SelectedItem.ToString().Split(' ')[1]);

                    //DEBUG
                    Console.WriteLine("comboBoxSerialParity.SelectedIndex:" + comboBoxSerialParity.SelectedIndex.ToString());

                    switch (comboBoxSerialParity.SelectedIndex)
                    {
                        case 0:
                            serialPort.Parity = Parity.None;
                            break;
                        case 1:
                            serialPort.Parity = Parity.Even;
                            break;
                        case 2:
                            serialPort.Parity = Parity.Odd;
                            break;
                        default:
                            serialPort.Parity = Parity.None;
                            break;
                    }

                    serialPort.DataBits = 8;

                    //DEBUG
                    Console.WriteLine("comboBoxSerialStop.SelectedIndex:" + comboBoxSerialStop.SelectedIndex.ToString());

                    switch (comboBoxSerialStop.SelectedIndex)
                    {
                        case 0:
                            serialPort.StopBits = StopBits.One;
                            break;
                        case 1:
                            serialPort.StopBits = StopBits.OnePointFive;
                            break;
                        case 2:
                            serialPort.StopBits = StopBits.Two;
                            break;
                        default:
                            serialPort.StopBits = StopBits.One;
                            break;
                    }

                    serialPort.Handshake = Handshake.None;

                    //Timeout porta
                    serialPort.ReadTimeout = 50;
                    serialPort.WriteTimeout = 50;

                    ModBus = new ModBus_Chicco(this, "RTU", (bool)CheckBoxSwapCrcBytes.IsChecked);

                    //Svuoto il buffer
                    //serialPort.DiscardInBuffer();
                    //serialPort.DiscardOutBuffer();

                    serialPort.Open();
                    serialPort.DiscardInBuffer();
                    serialPort.DiscardOutBuffer();

                    richTextBoxAppend(richTextBoxStatus, "Connected to " + comboBoxSerialPort.SelectedItem.ToString());

                    stringSerialPort = comboBoxSerialPort.SelectedItem.ToString();

                    loopSerial = new Thread(new ThreadStart(ModBus.loopSerialListenerRTU));
                    loopSerial.IsBackground = true;
                    loopSerial.Start();

                    radioButtonModeSerial.IsEnabled = false;
                    radioButtonModeTcp.IsEnabled = false;
                    checkBoxCheckModbusAddress.IsEnabled = false;

                    comboBoxSerialPort.IsEnabled = false;
                    comboBoxSerialSpeed.IsEnabled = false;
                    comboBoxSerialParity.IsEnabled = false;
                    comboBoxSerialStop.IsEnabled = false;

                    updateStartPauseStop();
                }
                catch (Exception error)
                {
                    pictureBoxSerial.Background = Brushes.LightGray;
                    pictureBoxRunningAs.Background = Brushes.LightGray;

                    try
                    {
                        //Potrebbe andare in crash se fallisce il try superiore
                        loopSerial.Abort();
                    }
                    catch { }

                    ModBus.threadTxRxIsRunning = false;

                    buttonSerialActive.Content = "Connect";

                    Console.WriteLine("Errore apertura porta seriale");
                    Console.WriteLine(error);

                    richTextBoxAppend(richTextBoxStatus, "Failed to connect");

                    comboBoxSerialPort.IsEnabled = true;
                    comboBoxSerialSpeed.IsEnabled = true;
                    comboBoxSerialParity.IsEnabled = true;
                    comboBoxSerialStop.IsEnabled = true;

                }
            }
            else
            {
                //Disattivazione comunicazione seriale
                pictureBoxSerial.Background = Brushes.LightGray;
                pictureBoxRunningAs.Background = Brushes.LightGray;

                loopSerial.Abort();

                ModBus.threadTxRxIsRunning = false;

                buttonSerialActive.Content = "Connect";

                radioButtonModeSerial.IsEnabled = true;
                radioButtonModeTcp.IsEnabled = true;
                checkBoxCheckModbusAddress.IsEnabled = true;

                comboBoxSerialPort.IsEnabled = true;
                comboBoxSerialSpeed.IsEnabled = true;
                comboBoxSerialParity.IsEnabled = true;
                comboBoxSerialStop.IsEnabled = true;

                //---------------------------------------------------------------------------------
                //----------------------Chiusura comunicazione seriale-----------------------------
                //---------------------------------------------------------------------------------
                serialPort.Close();
                richTextBoxAppend(richTextBoxStatus, "Port closed");
            }
        }

        private void buttonUpdateSerialList_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string[] SerialPortList = System.IO.Ports.SerialPort.GetPortNames();
                //comboBoxSerialPort.Items.Add("Seleziona porta seriale ...");
                comboBoxSerialPort.Items.Clear();

                foreach (string port in SerialPortList)
                {
                    comboBoxSerialPort.Items.Add(port);
                }
                comboBoxSerialPort.SelectedIndex = 0;
            }
            catch
            {
                Console.WriteLine("Nessuna porta seriale disponibile");
            }
        }

        // Visualizza console programma da menu tendina
        private void apriConsoleToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            apriConsole(); 
        }

        // Nasconde console programma da menu tendina
        private void chiudiConsoleToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            chiudiConsole();
        }

        public void chiudiConsole()
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);

            statoConsole = false;
        }

        public void apriConsole()
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_SHOW);

            statoConsole = true;
        }

        // Disabilita il pulsante di chiusura della console
        public void disableConsoleExitButton()
        {
            IntPtr handle = GetConsoleWindow();
            IntPtr exitButton = GetSystemMenu(handle, false);
            if (exitButton != null) DeleteMenu(exitButton, SC_CLOSE, MF_BYCOMMAND);
        }

        //----------------------------------------------------------------------------------
        //---------------------------SALVATAGGIO CONFIGURAZIONE-----------------------------
        //----------------------------------------------------------------------------------

        private void salva_configurazione(bool alert)   //Se alert true visualizza un messaggio di info salvataggio avvenuto
        {
            //DEBUG
            //MessageBox.Show("Salvataggio configurazione");

            try
            {
                // Caricamento variabili
                SAVE config = new SAVE();

                config.modbusAddress = textBoxModbusAddress.Text;
                config.usingSerial = (bool)radioButtonModeSerial.IsChecked;

                config.statoConsole = statoConsole;

                //config.serialMaster = radioButtonSerialMaster.IsChecked;
                config.serialRTU = true;

                //Serial port
                config.serialPort = comboBoxSerialPort.SelectedIndex;
                config.serialSpeed = comboBoxSerialSpeed.SelectedIndex;
                config.serialParity = comboBoxSerialParity.SelectedIndex;
                config.serialStop = comboBoxSerialStop.SelectedIndex;

                //TCP
                config.tcpClientIpAddress = textBoxTcpClientIpAddress.Text;
                config.tcpClientPort = textBoxTcpClientPort.Text;
                //config.tcpServerIpAddress = textBoxTcpServerIpAddress.Text;
                //config.tcpServerPort = textBoxTcpServerPort.Text;

                config.checkBoxUseOffsetInTables_ = true;
                config.checkBoxSavePackets_ = (bool)checkBoxSavePackets.IsChecked;
                config.checkBoxCloseConsolAfterBoot_ = (bool)checkBoxCloseConsolAfterBoot.IsChecked;

                config.textBoxSaveLogPath_ = textBoxSaveLogPath.Text;


                config.comboBoxCoilsRegisters_ = comboBoxCoilsRegisters.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsOffset_ = comboBoxCoilsOffset.SelectedValue.ToString().Split(' ')[1];
                config.textBoxCoilsOffset_ = textBoxCoilsOffset.Text;


                config.comboBoxInputsRegisters_ = comboBoxInputsRegisters.SelectedItem.ToString().Split(' ')[1];
                config.comboBoxInputsOffset_ = comboBoxInputsOffset.SelectedItem.ToString().Split(' ')[1];
                config.textBoxInputsOffset_ = textBoxInputsOffset.Text;

                config.comboBoxInputsRegRegisters_ = comboBoxInputsRegRegisters.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputsRegValues_ = comboBoxInputsRegValues.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputsRegOffset_ = comboBoxInputsRegOffset.SelectedValue.ToString().Split(' ')[1];
                config.textBoxInputsRegOffset_ = textBoxInputsRegOffset.Text;


                config.comboBoxHoldingsRegisters_ = comboBoxHoldingsRegisters.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingsValues_ = comboBoxHoldingsValues.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingsOffset_ = comboBoxHoldingsOffset.SelectedValue.ToString().Split(' ')[1];
                config.textBoxHoldingsOffset_ = textBoxHoldingsOffset.Text;

                config.checkBoxCheckModbusAddress_ = (bool)checkBoxCheckModbusAddress.IsChecked;

                config.colorDefaultReadCell_ = colorDefaultReadCell.ToString();
                config.colorDefaultWriteCell_ = colorDefaultWriteCell.ToString();
                config.colorErrorCell_ = colorErrorCell.ToString();

                config.checkBoxCorreggiRegistriModbus_ = (bool)checkBoxCorreggiRegistriModbus.IsChecked;
                config.nome_applicazione = textBoxNomeApplicazione.Text;

                config.checkBoxColorMenu_ = (bool)checkBoxColorMenu.IsChecked;
                config.colorMenuStrip_ = colorMenuStrip != null ? colorMenuStrip.ToString() : null;

                //---------------------------------------------
                //----------------Tabella coils----------------
                //---------------------------------------------

                string[] coils_A_ = new string[list_coilsTable.Count];
                string[] coils_B_ = new string[list_coilsTable.Count];
                string[] coils_C_ = new string[list_coilsTable.Count];

                for (int i = 0; i < (list_coilsTable.Count); i++)
                {
                    try
                    {
                        coils_A_[i] = list_coilsTable[i].Register;
                        coils_B_[i] = list_coilsTable[i].Value;
                        coils_C_[i] = list_coilsTable[i].Notes;
                    }
                    catch { }
                }

                config.coils_A = coils_A_;
                config.coils_B = coils_B_;
                config.coils_C = coils_C_;

                //---------------------------------------------
                //----------------Tabella inputs---------------
                //---------------------------------------------

                string[] inputs_A_ = new string[list_inputsTable.Count];
                string[] inputs_B_ = new string[list_inputsTable.Count];
                string[] inputs_C_ = new string[list_inputsTable.Count];

                for (int i = 0; i < (list_inputsTable.Count); i++)
                {
                    try
                    {
                        inputs_A_[i] = list_inputsTable[i].Register;
                        inputs_B_[i] = list_inputsTable[i].Value;
                        inputs_C_[i] = list_inputsTable[i].Notes;
                    }
                    catch { }
                }

                config.inputs_A = inputs_A_;
                config.inputs_B = inputs_B_;
                config.inputs_C = inputs_C_;

                //-------------------------------------------------------
                //----------------Tabella inputs registers---------------
                //-------------------------------------------------------

                string[] inputRegisters_A = new string[list_inputRegistersTable.Count];
                string[] inputRegisters_B = new string[list_inputRegistersTable.Count];
                string[] inputRegisters_C = new string[list_inputRegistersTable.Count];

                for (int i = 0; i < (list_inputRegistersTable.Count); i++)
                {
                    try
                    {
                        inputRegisters_A[i] = list_inputRegistersTable[i].Register;
                        inputRegisters_B[i] = list_inputRegistersTable[i].Value;
                        inputRegisters_C[i] = list_inputRegistersTable[i].Notes;
                    }
                    catch { }
                }

                config.inputRegisters_A = inputRegisters_A;
                config.inputRegisters_B = inputRegisters_B;
                config.inputRegisters_C = inputRegisters_C;

                //-------------------------------------------------------
                //----------------Tabella holdings-----------------------
                //-------------------------------------------------------

                string[] holding_A = new string[list_holdingRegistersTable.Count];
                string[] holding_B = new string[list_holdingRegistersTable.Count];
                string[] holding_C = new string[list_holdingRegistersTable.Count];

                for (int i = 0; i < (list_holdingRegistersTable.Count); i++)
                {
                    try
                    {
                        holding_A[i] = list_holdingRegistersTable[i].Register;
                        holding_B[i] = list_holdingRegistersTable[i].Value;
                        holding_C[i] = list_holdingRegistersTable[i].Notes;
                    }
                    catch { }
                }

                config.holdings_A = holding_A;
                config.holdings_B = holding_B;
                config.holdings_C = holding_C;

                config.checkBoxPinWIndow_ = (bool)checkBoxPinWIndow.IsChecked;
                config.checkBoxFocusReadRows_ = (bool)checkBoxFocusReadRows.IsChecked;
                config.checkBoxFocusWriteRows_ = (bool)checkBoxFocusWriteRows.IsChecked;
                config.checkBoxDisableGraphics_ = (bool)checkBoxDisableGraphics.IsChecked;
                config.textBoxOffsetFocusTabelle_ = textBoxOffsetFocusTabelle.Text;

                JavaScriptSerializer jss = new JavaScriptSerializer();
                string file_content = jss.Serialize(config);

                File.WriteAllText("Json/" + pathToConfiguration + "/CONFIGURAZIONE.json", file_content);

                if (alert)
                {
                    MessageBox.Show("Salvataggio configurazione avvenuto. Al prossimo avvio verrà caricata la configurazione corrente.", "Info");
                }

                Console.WriteLine("Salvata configurazione");
            }
            catch(Exception err)
            {
                Console.WriteLine("Errore salvataggio configurazione");
                Console.WriteLine(err);
            }
        }

        //----------------------------------------------------------------------------------
        //---------------------------CARICAMENTO CONFIGURAZIONE-----------------------------
        //----------------------------------------------------------------------------------

        private void carica_configurazione()
        {
            try
            {
                string file_content = File.ReadAllText("Json/" + pathToConfiguration + "/CONFIGURAZIONE.json");

                JavaScriptSerializer jss = new JavaScriptSerializer();
                SAVE config = jss.Deserialize<SAVE>(file_content);

                textBoxModbusAddress.Text = config.modbusAddress;
                radioButtonModeSerial.IsChecked = config.usingSerial;

                statoConsole = config.statoConsole;

                //Scheda configurazione seriale
                radioButtonModeSerial.IsChecked = config.usingSerial;
                radioButtonModeTcp.IsChecked = !config.usingSerial;

                //radioButtonSerialMaster.IsChecked = config.serialMaster;
                //radioButtonSerialSlave.IsChecked = !config.serialMaster;

                comboBoxSerialSpeed.SelectedIndex = config.serialSpeed;
                comboBoxSerialParity.SelectedIndex = config.serialParity;
                comboBoxSerialStop.SelectedIndex = config.serialStop;

                //Scheda configurazione TCP
                //radioButtonTcpMaster.IsChecked = config.serialMaster;
                //radioButtonTcpSlave.IsChecked = !config.serialMaster;

                textBoxTcpClientIpAddress.Text = config.tcpClientIpAddress;
                textBoxTcpClientPort.Text = config.tcpClientPort;
                //textBoxTcpServerIpAddress.Text = config.tcpServerIpAddress;
                //textBoxTcpServerPort.Text = config.tcpServerPort;


                //comboBoxSerialPort.SelectedIndex = config.serialPort;

                checkBoxSavePackets.IsChecked = config.checkBoxSavePackets_;
                checkBoxCloseConsolAfterBoot.IsChecked = config.checkBoxCloseConsolAfterBoot_;


                textBoxSaveLogPath.Text = config.textBoxSaveLogPath_;

                comboBoxCoilsRegisters.SelectedIndex = config.comboBoxCoilsRegisters_.ToString() == "HEX" ? 1 : 0;
                comboBoxCoilsOffset.SelectedItem = config.comboBoxCoilsOffset_.ToString() == "HEX" ? 1 : 0;
                textBoxCoilsOffset.Text = config.textBoxCoilsOffset_;

                comboBoxInputsRegisters.SelectedItem = config.comboBoxInputsRegisters_.ToString() == "HEX" ? 1 : 0;
                comboBoxInputsOffset.SelectedItem = config.comboBoxInputsOffset_.ToString() == "HEX" ? 1 : 0;
                textBoxInputsOffset.Text = config.textBoxInputsOffset_;

                comboBoxInputsRegRegisters.SelectedItem = config.comboBoxInputsRegRegisters_.ToString() == "HEX" ? 1 : 0;
                comboBoxInputsRegValues.SelectedItem = config.comboBoxInputsRegValues_.ToString() == "HEX" ? 1 : 0;
                comboBoxInputsRegOffset.SelectedItem = config.comboBoxInputsRegOffset_.ToString() == "HEX" ? 1 : 0;
                textBoxInputsRegOffset.Text = config.textBoxInputsRegOffset_;

                comboBoxHoldingsRegisters.SelectedItem = config.comboBoxHoldingsRegisters_.ToString() == "HEX" ? 1 : 0;
                comboBoxHoldingsValues.SelectedItem = config.comboBoxHoldingsValues_.ToString() == "HEX" ? 1 : 0;
                comboBoxHoldingsOffset.SelectedItem = config.comboBoxHoldingsOffset_.ToString() == "HEX" ? 1 : 0;
                textBoxHoldingsOffset.Text = config.textBoxHoldingsOffset_;

                checkBoxCheckModbusAddress.IsChecked = config.checkBoxCheckModbusAddress_;

                BrushConverter bc = new BrushConverter();
                
                colorDefaultReadCell = (SolidColorBrush)bc.ConvertFromString(config.colorDefaultReadCell_);
                colorDefaultWriteCell = (SolidColorBrush)bc.ConvertFromString(config.colorDefaultWriteCell_);
                colorErrorCell = (SolidColorBrush)bc.ConvertFromString(config.colorErrorCell_);

                labelColorCellRead.Background = colorDefaultReadCell;
                labelColorCellWrote.Background = colorDefaultWriteCell;
                labelColorCellError.Background = colorErrorCell;

                checkBoxColorMenu.IsChecked = config.checkBoxColorMenu_;
                if (config.colorMenuStrip_ != null) { colorMenuStrip = (SolidColorBrush)bc.ConvertFromString(config.colorMenuStrip_); }

                if ((bool)checkBoxColorMenu.IsChecked)
                {
                    menuStrip.Background = colorMenuStrip;
                }

                checkBoxCorreggiRegistriModbus.IsChecked = config.checkBoxCorreggiRegistriModbus_;
                textBoxNomeApplicazione.Text = config.nome_applicazione;

                //Caricamento tabelle

                //---------------------------------------------
                //----------------Tabella coils----------------
                //---------------------------------------------

                String[] a = config.coils_A;
                String[] b = config.coils_B;
                String[] c = config.coils_C;

                list_coilsTable.Clear();

                for (int i = 0; i < (a.Length); i++)
                {
                    try
                    {
                        ModBus_Item  row = new ModBus_Item();

                        row.Register = a[i];
                        row.Value = b[i];
                        row.Notes = c[i];

                        list_coilsTable.Add(row);
                    }
                    catch { }
                }

                //---------------------------------------------
                //----------------Tabella inputs----------------
                //---------------------------------------------

                String[] inputs_A = config.inputs_A;
                String[] inputs_B = config.inputs_B;
                String[] inputs_C = config.inputs_C;

                list_inputsTable.Clear();

                for (int i = 0; i < (inputs_A.Length); i++)
                {
                    try
                    {
                        ModBus_Item row = new ModBus_Item();

                        row.Register = inputs_A[i];
                        row.Value = inputs_B[i];
                        row.Notes = inputs_C[i];

                        list_inputsTable.Add(row);
                    }
                    catch { }
                }

                //-------------------------------------------------------
                //----------------Tabella input registers----------------
                //-------------------------------------------------------

                String[] inputRegisters_A = config.inputRegisters_A;
                String[] inputRegisters_B = config.inputRegisters_B;
                String[] inputRegisters_C = config.inputRegisters_C;

                list_inputRegistersTable.Clear();

                for (int i = 0; i < (inputRegisters_A.Length); i++)
                {
                    try
                    {
                        ModBus_Item row = new ModBus_Item();

                        row.Register = inputRegisters_A[i];
                        row.Value = inputRegisters_B[i];
                        row.Notes = inputRegisters_C[i];

                        list_inputRegistersTable.Add(row);
                    }
                    catch { }
                }

                //-------------------------------------------------------
                //--------------------Tabella holdings-------------------
                //-------------------------------------------------------

                String[] holdings_A = config.holdings_A;
                String[] holdings_B = config.holdings_B;
                String[] holdings_C = config.holdings_C;

                list_holdingRegistersTable.Clear();

                for (int i = 0; i < (holdings_A.Length); i++)
                {
                    try
                    {
                        ModBus_Item row = new ModBus_Item();

                        row.Register = holdings_A[i];
                        row.Value = holdings_B[i];
                        row.Notes = holdings_C[i];

                        list_holdingRegistersTable.Add(row);
                    }
                    catch { }
                }

                //Show/hide console
                if (!statoConsole)
                {
                    var handle = GetConsoleWindow();
                    ShowWindow(handle, SW_HIDE);
                }

                // Sulle ultime variabili aggiunte faccio un controllo se esistono nel file di configurazione, su file vecchi potrebbero non esistere
                if (config.checkBoxPinWIndow_ != null)
                {
                    checkBoxPinWIndow.IsChecked = config.checkBoxPinWIndow_;
                }
                
                if(config.checkBoxFocusReadRows_ != null) 
                {
                    checkBoxFocusReadRows.IsChecked = config.checkBoxFocusReadRows_;
                }

                if (config.checkBoxFocusWriteRows_ != null)
                {
                    checkBoxFocusWriteRows.IsChecked = config.checkBoxFocusWriteRows_;
                }

                if (config.checkBoxDisableGraphics_ != null)
                {
                    checkBoxDisableGraphics.IsChecked = config.checkBoxDisableGraphics_;
                }

                if (config.textBoxOffsetFocusTabelle_ != null)
                {
                    textBoxOffsetFocusTabelle.Text = config.textBoxOffsetFocusTabelle_;
                }

                Console.WriteLine("Caricata configurazione precedente\n");
            }
            catch (Exception error)
            {
                Console.WriteLine("Errore caricamento configurazione\n");
                Console.WriteLine(error);
            }
        }

        private void salvaToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            salva_configurazione(true);
        }



        private void buttonTcpActive_Click(object sender, RoutedEventArgs e)
        {
            String ip_address = "";
            String port = "";

            try
            {
                if (pictureBoxTcp.Background == Brushes.LightGray)
                {
                    ip_address = textBoxTcpClientIpAddress.Text;
                    port = textBoxTcpClientPort.Text;

                    //Tcp attivo
                    pictureBoxTcp.Background = Brushes.Lime;
                    pictureBoxRunningAs.Background = Brushes.Lime;

                    ModBus = new ModBus_Chicco(this, "TCP", (bool)CheckBoxSwapCrcBytes.IsChecked);

                    richTextBoxAppend(richTextBoxStatus, "Server active at " + ip_address + ":" + port);

                    loopTcp = new Thread(new ThreadStart(ModBus.loopTcpListener));
                    loopTcp.Priority = ThreadPriority.Highest;
                    loopTcp.IsBackground = false;
                    loopTcp.Start();

                    radioButtonModeSerial.IsEnabled = false;
                    radioButtonModeTcp.IsEnabled = false;
                    //checkBoxCheckModbusAddress.IsEnabled = false;

                    //textBoxModbusAddress.IsEnabled = false;

                    buttonTcpActive.Content = "Stop";

                    textBoxTcpClientIpAddress.IsEnabled = false;
                    textBoxTcpClientPort.IsEnabled = false;

                    updateStartPauseStop();
                }
                else
                {
                    //Tcp disattivo
                    pictureBoxTcp.Background = Brushes.LightGray;
                    pictureBoxRunningAs.Background = Brushes.LightGray;

                    richTextBoxAppend(richTextBoxStatus, "Server stopped");

                    // Essendo la funzione TCP bloccante devo fermare il server prima di abortire il thread
                    ModBus.stopTcpLoop();

                    // Chiudo il thread del loop TCP
                    loopTcp.Abort();

                    // MANCA PARTE SULLO STOP DEL THREAD

                    buttonTcpActive.Content = "Start";

                    radioButtonModeSerial.IsEnabled = true;
                    radioButtonModeTcp.IsEnabled = true;
                    checkBoxCheckModbusAddress.IsEnabled = true;

                    textBoxTcpClientIpAddress.IsEnabled = true;
                    textBoxTcpClientPort.IsEnabled = true;
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        // Altri pulsanti nella grafica
        private void guidaToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(Directory.GetCurrentDirectory() + "\\Manuali\\Guida_ModBus_Server_" + textBoxCurrentLanguage.Text + ".pdf");
            }
            catch
            {
                MessageBox.Show("Ancora da scrivere :-)", "Hey");
            }
        }

        private void infoToolStripMenuItem1_Click(object sender, RoutedEventArgs e)
        {
            Info info = new Info(title, version);

            info.Show();
        }

        private void buttonClearReceived_Click(object sender, RoutedEventArgs e)
        {
            richTextBoxPackets.Document.Blocks.Clear();
            richTextBoxPackets.AppendText("\n");
        }

        private void richTextBoxAppend(RichTextBox richTextBox, String append)
        {
            richTextBox.AppendText(DateTime.Now.ToString().Split(' ')[1] + " - " + append + "\n");
            richTextBox.ScrollToEnd();
        }

        private void buttonClearSerialStatus_Click(object sender, RoutedEventArgs e)
        {
            count_status = 0;

            richTextBoxStatus.Document.Blocks.Clear();
            richTextBoxStatus.AppendText("\n");
        }

        private void buttonClearTcpStatus_Click(object sender, RoutedEventArgs e)
        {
            richTextBoxStatus.Document.Blocks.Clear();
            richTextBoxStatus.AppendText("\n");
        }

        private void buttonSaveLogSending_Click(object sender, RoutedEventArgs e)
        {
            //saveFileDialogBox.DefaultExt = "txt";
            //saveFileDialogBox.ShowDialog();

            MessageBox.Show("Funzione non ancora implementata", "Info");
        }

        private void esciSenzaSalvareToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void esciToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            salva_configurazione(false);
            this.Close();
        }

        private void impostazioniToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            tabControlMain.SelectedIndex = 7;
        }

        // Salvataggio pacchetti
        private void buttonExportSentPackets_Click(object sender, RoutedEventArgs e)
        {
            saveFileDialogBox = new SaveFileDialog();

            saveFileDialogBox.DefaultExt = ".txt";
            saveFileDialogBox.AddExtension = false;
            saveFileDialogBox.FileName = "Pacchetti_inviati_" + DateTime.Now.Day.ToString().PadLeft(2, '0') + "." + DateTime.Now.Month.ToString().PadLeft(2, '0') + "." + DateTime.Now.Year.ToString().PadLeft(2, '0');
            saveFileDialogBox.Filter = "Text|*.txt|Log|*.log";
            saveFileDialogBox.Title = "Salva Log pacchetti inviati";

            if ((bool)saveFileDialogBox.ShowDialog())
            {
                StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                TextRange textRange = new TextRange(
                    // TextPointer to the start of content in the RichTextBox.
                    richTextBoxPackets.Document.ContentStart,
                    // TextPointer to the end of content in the RichTextBox.
                    richTextBoxPackets.Document.ContentEnd
                );

                writer.Write(textRange.Text);
                writer.Dispose();
                writer.Close();
            }
        }

        private void coilsTableImportToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            importDataGrid(list_coilsTable, "", "Carica tabella coils", comboBoxCoilsOffset, textBoxCoilsOffset, comboBoxCoilsRegisters);
        }

        private void coilsExportToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            string offset = textBoxCoilsOffset.Text;

            if (comboBoxCoilsOffset.SelectedIndex == 1)
            {
                offset = "0x" + offset;
            }

            exportDataGrid(list_coilsTable, "Coils_", "Salva tabella coils", offset, comboBoxCoilsOffset.SelectedIndex == 1);
        }

        public void exportDataGrid(ObservableCollection<ModBus_Item> collection, String fileName, String title, String offset, bool registerHex)
        {
            try
            {
                saveFileDialogBox = new SaveFileDialog();

                saveFileDialogBox.DefaultExt = "csv";
                saveFileDialogBox.AddExtension = false;
                saveFileDialogBox.FileName = "" + fileName;
                //saveFileDialogBox.Filter = "CSV|*.csv|JSON|*.json";
                saveFileDialogBox.Filter = "CSV|*.csv";
                saveFileDialogBox.Title = title;

                if ((bool)saveFileDialogBox.ShowDialog())
                {
                    String content = "Offset,Register,Value,Notes,Mappings\n";

                    foreach (ModBus_Item item in collection)
                    {
                        if (item != null)
                        {
                            if (registerHex)
                            {
                                content += offset + ",0x" + item.Register + "," + item.Value + "," + item.Notes + "," + "\n";
                            }
                            else
                            {
                                content += offset + "," + item.Register + "," + item.Value + "," + item.Notes + "," + "\n";
                            }
                        }
                    }

                    StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                    writer.Write(content);
                    writer.Dispose();
                    writer.Close();

                    /*if (saveFileDialogBox.FileName.IndexOf("csv") != 1)
                    {
                        
                    }
                    else
                    {
                        // File JSON
                        dataGridJson save = new dataGridJson();

                        String[] a = new string[(dataGrid.Count)];
                        String[] b = new string[(dataGrid.Count)];
                        String[] c = new string[(dataGrid.Count)];

                        for (int i = 0; i < (dataGrid.Count); i++)
                        {
                            try
                            {
                                a[i] = dataGrid[i].Register;
                                b[i] = dataGrid[i].Value;
                                c[i] = dataGrid[i].Notes;
                            }
                            catch { }
                        }

                        save.registers = a;
                        save.values = b;
                        save.note = c;

                        JavaScriptSerializer jss = new JavaScriptSerializer();
                        string file_content = jss.Serialize(save);

                        StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                        writer.Write(file_content);
                        writer.Dispose();
                        writer.Close();
                    }*/
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        public void importDataGrid(ObservableCollection<ModBus_Item> collection, String fileName, String title, ComboBox comboBoxOffset, TextBox textBoxOffset, ComboBox comboBoxReg)
        {
            try
            {
                openFileDialogBox = new OpenFileDialog();

                openFileDialogBox.DefaultExt = "csv|json";
                openFileDialogBox.AddExtension = false;
                openFileDialogBox.FileName = "" + fileName;
                openFileDialogBox.Filter = "CSV|*.csv|JSON|*.json";
                openFileDialogBox.Title = title;
            
                if ((bool)openFileDialogBox.ShowDialog())
                {
                    //DEBUG
                    Console.WriteLine(openFileDialogBox.FileName);

                    // File CSV

                    collection.Clear();

                    string content = File.ReadAllText(openFileDialogBox.FileName);
                    string[] splitted = content.Split('\n');

                    for (int i = 1; i < splitted.Count(); i++)
                    {
                        ModBus_Item item = new ModBus_Item();

                        try
                        {
                            if (splitted[i].Length > 2)
                            {
                                string offset = splitted[i].Split(',')[0];

                                if (offset.IndexOf("0x") != -1)
                                {
                                    comboBoxOffset.SelectedIndex = 1;
                                    offset = offset.Substring(2);
                                }

                                textBoxOffset.Text = offset;

                                item.Register = splitted[i].Split(',')[1];

                                if (item.Register.IndexOf("0x") != -1)
                                {
                                    comboBoxReg.SelectedIndex = 1;
                                    item.Register = item.Register.Substring(2);
                                }

                                item.Value = splitted[i].Split(',')[2];

                                if (item.Value.IndexOf("0x") != -1)
                                {
                                    comboBoxReg.SelectedIndex = 1;
                                    item.Value = item.Register.Substring(2);
                                }

                                item.Notes = splitted[i].Split(',')[3];
                                //item.Mappings = splitted[i].Split(',')[4];

                                collection.Add(item);
                            }
                        }
                        catch
                        {
                            //Console.WriteLine(err);
                        }
                    }

                    /*if (openFileDialogBox.FileName.IndexOf(".csv") == -1)
                    {
                        // File JSON

                        string file_content = File.ReadAllText(openFileDialogBox.FileName);

                        JavaScriptSerializer jss = new JavaScriptSerializer();
                        dataGridJson load = jss.Deserialize<dataGridJson>(file_content);

                        String[] a = load.registers;
                        String[] b = load.values;
                        String[] c = load.note;

                        dataGrid.Clear();

                        for (int i = 0; i < (a.Length); i++)
                        {
                            try
                            {
                                ModBus_Item row = new ModBus_Item();

                                row.Register = a[i];
                                row.Value = b[i];
                                row.Notes = c[i];

                                dataGrid.Add(row);
                            }
                            catch { }
                        }
                    }

                    else
                    {
                        
                    }*/
                }

            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void inputExportToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            string offset = textBoxInputsOffset.Text;

            if (comboBoxInputsOffset.SelectedIndex == 1)
            {
                offset = "0x" + offset;
            }

            exportDataGrid(list_inputsTable, "Inputs_", "Salva tabella inputs", offset, comboBoxInputsOffset.SelectedIndex == 1);
        }

        private void inputRegisterExportToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            string offset = textBoxInputsRegOffset.Text;

            if (comboBoxInputsRegOffset.SelectedIndex == 1)
            {
                offset = "0x" + offset;
            }

            exportDataGrid(list_inputRegistersTable, "Input_registers_", "Salva tabella input registers", offset, comboBoxInputsRegOffset.SelectedIndex == 1);
        }

        private void holdingExportToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            string offset = textBoxHoldingsOffset.Text;

            if (comboBoxHoldingsOffset.SelectedIndex == 1)
            {
                offset = "0x" + offset;
            }

            exportDataGrid(list_holdingRegistersTable, "Holdings_", "Salva tabella holdings", offset, comboBoxHoldingsOffset.SelectedIndex == 1);
        }

        private void inputImportToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            importDataGrid(list_inputsTable, "", "Carica tabella inputs", comboBoxInputsOffset, textBoxInputsOffset, comboBoxInputsRegisters);
        }

        private void inputRegisterImportToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            importDataGrid(list_inputRegistersTable, "", "Carica tabella input registers", comboBoxInputsRegOffset, textBoxInputsRegOffset, comboBoxInputsRegRegisters);
        }

        private void holdingImportToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            importDataGrid(list_holdingRegistersTable, "", "Carica tabella holdings", comboBoxHoldingsOffset, textBoxHoldingsOffset, comboBoxHoldingsRegisters);
        }

        private void buttonColorCellRead_Click_1(object sender, RoutedEventArgs e)
        {
            if (colorDialogBox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                colorDefaultReadCell = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));
                labelColorCellRead.Background = colorDefaultReadCell;
            }
        }

        private void buttonCellWrote_Click_1(object sender, RoutedEventArgs e)
        {
            if (colorDialogBox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                colorDefaultWriteCell = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));
                labelColorCellWrote.Background = colorDefaultWriteCell;
            }
        }

        private void buttonColorCellError_Click_1(object sender, RoutedEventArgs e)
        {
            if (colorDialogBox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                colorErrorCell = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));
                labelColorCellError.Background = colorErrorCell;
            }
        }

        private void textBoxNomeApplicazione_TextChanged(object sender, RoutedEventArgs e)
        {
            this.Title = "ModBus C# Server " + version + " - " + textBoxNomeApplicazione.Text;
        }

        private void checkBoxCorreggiRegistriModbus_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if ((bool)checkBoxCorreggiRegistriModbus.IsChecked)
            {
                comboBoxCoilsRegisters.IsEnabled = false;
                comboBoxInputsRegisters.IsEnabled = false;
                comboBoxInputsRegRegisters.IsEnabled = false;
                comboBoxHoldingsRegisters.IsEnabled = false;

                comboBoxCoilsRegisters.SelectedItem = "DEC";
                comboBoxInputsRegisters.SelectedItem = "DEC";
                comboBoxInputsRegRegisters.SelectedItem = "DEC";
                comboBoxHoldingsRegisters.SelectedItem = "DEC";

                labelPrimoRegistroCoils.Content = "Range registri: 1-9999";
                labelPrimoRegistroInputs.Content = "Range registri: 10001-19999";
                labelPrimoInputRegister.Content = "Range registri: 30001-39999";
                labelPrimoRegistroHolding.Content = "Range registri: 40001-49999";
            }
            else
            {
                comboBoxCoilsRegisters.IsEnabled = true;
                comboBoxInputsRegisters.IsEnabled = true;
                comboBoxInputsRegRegisters.IsEnabled = true;
                comboBoxHoldingsRegisters.IsEnabled = true;

                labelPrimoRegistroCoils.Content = "Range registri: 0-9998";
                labelPrimoRegistroInputs.Content = "Range registri: 0-9998";
                labelPrimoInputRegister.Content = "Range registri: 0-9998";
                labelPrimoRegistroHolding.Content = "Range registri: 0-9998";
            }
        }

        private void buttonColorMenu_Click(object sender, RoutedEventArgs e)
        {
            /*if (colorDialogBox.ShowDialog() == DialogResult.OK)
            {
                colorMenuStrip = colorDialogBox.Color;
                menuStrip1.Background = colorMenuStrip;
            }*/
        }

        private void checkBoxColorMenu_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if ((bool)checkBoxColorMenu.IsChecked)
            {
                menuStrip.Background = colorMenuStrip;
            }
            else
            {
                menuStrip.Background = SystemColors.ControlBrush;
            }
        }

        private void buttonClearAll_Click(object sender, RoutedEventArgs e)
        {
            count_log = 0;

            richTextBoxPackets.Document.Blocks.Clear();
            richTextBoxPackets.AppendText("\n");
        }

        private void checkBoxAddLinesToEnd_CheckedChanged(object sender, RoutedEventArgs e)
        {
            //ModBus.insertLogLinesAtTop(!(bool)checkBoxAddLinesToEnd.IsChecked);
        }

        private void salvaConfigurazioneNelDatabaseToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            String prevoiusPath = pathToConfiguration;

            Salva_impianto form_save = new Salva_impianto();
            form_save.ShowDialog();

            //Controllo il risultato del form
            if ((bool)form_save.DialogResult)
            {
                pathToConfiguration = form_save.path;

                salva_configurazione(false);

                if (pathToConfiguration != defaultPathToConfiguration)
                {
                    this.Title = "ModBus C# Client " + version + " - File: " + pathToConfiguration;
                }


                Directory.CreateDirectory("Json\\" + pathToConfiguration);

                String[] fileNames = Directory.GetFiles("Json\\" + prevoiusPath + "\\");

                for (int i = 0; i < fileNames.Length; i++)
                {
                    String newFile = "Json\\" + pathToConfiguration + fileNames[i].Substring(fileNames[i].LastIndexOf('\\'));

                    Console.WriteLine("Copying file: " + fileNames[i] + " to " + newFile);
                    File.Copy(fileNames[i], newFile);
                }
            }
        }

        private void caricaConfigurazioneDalDatabaseToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            Carica_impianto form_load = new Carica_impianto(defaultPathToConfiguration, this);
            form_load.ShowDialog();

            //Controllo il risultato del form
            if ((bool)form_load.DialogResult)
            {
                salva_configurazione(false);

                pathToConfiguration = form_load.path;

                // debug
                //Console.WriteLine("pathToConfiguration: " + pathToConfiguration);
                //Console.WriteLine("defaultPathToConfiguration: " + defaultPathToConfiguration);

                if (pathToConfiguration != defaultPathToConfiguration)
                {
                    this.Title = "ModBus C# Server " + version + " - File: " + pathToConfiguration;
                }

                carica_configurazione();
            }
        }

        private void gestisciDatabaseToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("explorer.exe", "Json");
            }
            catch
            {
                MessageBox.Show("Imposssibile aprire la cartella di configurazione Database", "Alert");
                Console.WriteLine("Imposssibile aprire la cartella di configurazione Database");
            }
        }

        private void buttonPingIp_Click(object sender, RoutedEventArgs e)
        {
            Ping p1 = new Ping();
            PingReply PR = p1.Send(textBoxTcpClientIpAddress.Text, 500);

            // check when the ping is not success
            if (!PR.Status.ToString().Equals("Success"))
            {
                buttonPingIp.Background = Brushes.Red;
                DoEvents();
                MessageBox.Show("Ping failed", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                buttonPingIp.Background = Brushes.LightGreen;
                DoEvents();
                MessageBox.Show("Ping ok.\nResponse time: " + PR.RoundtripTime + "ms", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void licenseToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            License window = new License();
            window.Show();
        }

        private void CheckBoxPinWIndow_Checked(object sender, RoutedEventArgs e)
        {
            this.Topmost = (bool)checkBoxPinWIndow.IsChecked;
        }

        private void ButtonStart_Click(object sender, RoutedEventArgs e)
        {
            // Avvia server (risponde alle query con il contenuto delle tabelle)
            if (ModBus != null)
            {
                ModBus.mode = 0;

                updateStartPauseStop();
            }
        }

        private void ButtonPause_Click(object sender, RoutedEventArgs e)
        {
            // Pausa server (risponde alle query senza aggiornare le tabelle, per editing senza far crashare l'applicazione)
            if (ModBus != null)
            {
                ModBus.mode = 1;

                updateStartPauseStop();
            }
        }

        private void ButtonStopLog_Click(object sender, RoutedEventArgs e)
        {
            // Stop server (smette di rispondere alle richieste)
            if (ModBus != null)
            {
                ModBus.mode = 2;

                updateStartPauseStop();
            }
        }

        public void updateStartPauseStop()
        {
            // IconPlay.Brush = new SolidColorBrush(Color.FromArgb(0xFF, 0x00, 0xF0, 0x00));
            // IconPause.Brush = new SolidColorBrush(Color.FromArgb(0xFF, 0x00, 0x53, 0x9C));
            // IconStop.Brush = new SolidColorBrush(Color.FromArgb(0xFF, 0xF0, 0x00, 0x00));

            if (ModBus != null)
            {
                switch (ModBus.mode)
                {
                    case 0: // Start
                        IconPlay.Brush = Brushes.LightGray;
                        IconPause.Brush = new SolidColorBrush(Color.FromArgb(0xFF, 0x00, 0x53, 0x9C));
                        IconStop.Brush = new SolidColorBrush(Color.FromArgb(0xFF, 0xF0, 0x00, 0x00));

                        LabelCurrentMode.Content = "Running";
                        LabelCurrentMode.Foreground = Brushes.LimeGreen;

                        richTextBoxStatus.AppendText(DateTime.Now.ToString().Split(' ')[1] + " - Running\n");

                        // Sull'evento di start aggiorno la tabella
                        dataGridViewCoils.ScrollIntoView(dataGridViewCoils.Items[0]);
                        dataGridViewCoils.UpdateLayout();
                        //dataGridViewCoils.Items.Refresh();

                        dataGridViewInput.ScrollIntoView(dataGridViewInput.Items[0]);
                        dataGridViewInput.UpdateLayout();
                        //dataGridViewInput.Items.Refresh();

                        dataGridViewInputRegister.ScrollIntoView(dataGridViewInputRegister.Items[0]);
                        dataGridViewInputRegister.UpdateLayout();
                        //dataGridViewInputRegister.Items.Refresh();

                        dataGridViewHolding.ScrollIntoView(dataGridViewHolding.Items[0]);
                        dataGridViewHolding.UpdateLayout();
                        //dataGridViewHolding.Items.Refresh();
                        break;

                    case 1: // Paused
                        IconPlay.Brush = new SolidColorBrush(Color.FromArgb(0xFF, 0x00, 0xF0, 0x00));
                        IconPause.Brush = Brushes.LightGray;
                        IconStop.Brush = new SolidColorBrush(Color.FromArgb(0xFF, 0xF0, 0x00, 0x00));

                        LabelCurrentMode.Content = "Edit mode";
                        LabelCurrentMode.Foreground = new SolidColorBrush(Color.FromArgb(0xFF, 0x00, 0x53, 0x9C)); ;

                        richTextBoxStatus.AppendText(DateTime.Now.ToString().Split(' ')[1] + " - Edit mode\n");
                        break;

                    case 2: // Stopped
                        IconPlay.Brush = new SolidColorBrush(Color.FromArgb(0xFF, 0x00, 0xF0, 0x00));
                        IconPause.Brush = new SolidColorBrush(Color.FromArgb(0xFF, 0x00, 0x53, 0x9C)); 
                        IconStop.Brush = Brushes.LightGray;

                        LabelCurrentMode.Content = "Stopped";
                        LabelCurrentMode.Foreground = new SolidColorBrush(Color.FromArgb(0xFF, 0xF0, 0x00, 0x00)); ;

                        richTextBoxStatus.AppendText(DateTime.Now.ToString().Split(' ')[1] + " - Stopped\n");
                        break;
                }
            }
            else
            {
                // Disabled
                IconPlay.Brush = Brushes.LightGray;
                IconPause.Brush = Brushes.LightGray;
                IconStop.Brush = Brushes.LightGray;

                LabelCurrentMode.Content = "";
                LabelCurrentMode.Background = Brushes.Transparent;
            }
        }

        private void checkBoxFocusReadRows_Checked(object sender, RoutedEventArgs e)
        {
            if (ModBus != null)
            {
                ModBus.scrollToReadRows = (bool)checkBoxFocusReadRows.IsChecked;
            }
        }

        private void checkBoxWriteRows_Checked(object sender, RoutedEventArgs e)
        {
            if (ModBus != null)
            {
                ModBus.scrollToWriteRows = (bool)checkBoxFocusWriteRows.IsChecked;
            }
        }

        private void checkBoxDisableGraphics_Checked(object sender, RoutedEventArgs e)
        {
            if (ModBus != null)
            {
                ModBus.disableGraphics = (bool)checkBoxDisableGraphics.IsChecked;
            }

            checkBoxFocusReadRows.IsEnabled = !(bool)checkBoxDisableGraphics.IsChecked;
            checkBoxFocusWriteRows.IsEnabled = !(bool)checkBoxDisableGraphics.IsChecked;
            dataGridViewCoils.IsEnabled = !(bool)checkBoxDisableGraphics.IsChecked;
            dataGridViewInput.IsEnabled = !(bool)checkBoxDisableGraphics.IsChecked;
            dataGridViewInputRegister.IsEnabled = !(bool)checkBoxDisableGraphics.IsChecked;
            dataGridViewHolding.IsEnabled = !(bool)checkBoxDisableGraphics.IsChecked;
        }

        private void textBoxOffsetFocusTabelle_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(ModBus != null)
            {
                if(!int.TryParse(textBoxOffsetFocusTabelle.Text, out ModBus.offsetFocusTabelle))
                {
                    ModBus.offsetFocusTabelle = 5;
                }
            }
        }

        public static void DoEvents()
        {
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, new Action(delegate { }));
        }

        private void dataGridViewHolding_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            if (ModBus != null)
            {
                if (ModBus.mode == 0)
                {
                    ModBus.mode = 1;
                    updateStartPauseStop();

                    if (!ShowAlertEditMode)
                    {
                        MessageBox.Show("Il programma è passato in modalità edit, continua a rispondere ai comandi ModBus con i valori nelle tabelle ma non aggiorna più i colori e non salta più alle righe lette/scritte dal master. Premere start per riavviare la grafica al termine delle modifiche.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                        ShowAlertEditMode = true;
                    }
                }
            }
        }

        private void dataGridViewInputRegister_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            if (ModBus != null)
            {
                if (ModBus.mode == 0)
                {
                    ModBus.mode = 1;
                    updateStartPauseStop();
                    MessageBox.Show("Il programma è passato in modalità edit, continua a rispondere ai comandi ModBus ma non aggiorna più la grafica delle tabelle. Premere start per riavviare la grafica al termine delle modifiche.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        private void dataGridViewInput_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            if (ModBus != null)
            {
                if (ModBus.mode == 0)
                {
                    ModBus.mode = 1;
                    updateStartPauseStop();
                    MessageBox.Show("Il programma è passato in modalità edit, continua a rispondere ai comandi ModBus ma non aggiorna più la grafica delle tabelle.Premere start per riavviare la grafica al termine delle modifiche.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        private void dataGridViewCoils_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            if (ModBus != null)
            {
                if (ModBus.mode == 0)
                {
                    ModBus.mode = 1;
                    updateStartPauseStop();
                    MessageBox.Show("Il rogramma è passato in modalità edit, contiunua a rispondere al ModBus ma non aggiorna più la grafica delle tabelle. Premere start per riavviare la grafica al termine delle modifiche.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        public void LogDequeue()
        {
            while (!dequeueExit)
            {
                String content;

                if (this.ModBus != null)
                {
                    if (this.ModBus.log2.TryDequeue(out content))
                    {
                        richTextBoxPackets.Dispatcher.Invoke((Action)delegate
                        {
                            if(count_log > LogLimitRichTextBox)
                            {
                                // Arrivato al limite tolgo una riga ogni volta che aggiungo una riga
                                richTextBoxPackets.Document.Blocks.Remove(richTextBoxPackets.Document.Blocks.FirstBlock);
                            }
                            else
                            {
                                count_log += 1;
                            }

                            richTextBoxPackets.AppendText(content);
                        });

                        scrolled_log = false;
                    }
                    else
                    {
                        if (!scrolled_log)
                        {
                            richTextBoxPackets.Dispatcher.Invoke((Action)delegate
                            {
                                richTextBoxPackets.ScrollToEnd();
                            });

                            scrolled_log = true;
                        }

                        Thread.Sleep(100);
                    }

                    if (this.ModBus.logStatus.TryDequeue(out content))
                    {
                        richTextBoxStatus.Dispatcher.Invoke((Action)delegate
                        {
                            if (count_status > LogLimitRichTextBox)
                            {
                                // Arrivato al limite tolgo una riga ogni volta che aggiungo una riga
                                richTextBoxStatus.Document.Blocks.Remove(richTextBoxStatus.Document.Blocks.FirstBlock);
                            }
                            else
                            {
                                count_status += 1;
                            }

                            richTextBoxStatus.AppendText(content);

                            scrolled_status = false;
                        });
                    }
                    else
                    {
                        if (!scrolled_status)
                        {
                            richTextBoxStatus.Dispatcher.Invoke((Action)delegate
                            {
                                richTextBoxStatus.ScrollToEnd();
                            });

                            scrolled_status = true;
                        }

                        Thread.Sleep(100);
                    }
                }
                else
                {
                    Thread.Sleep(100);
                }
            }
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {

            // Vincolato al ctrl
            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                switch (e.Key)
                {
                    case Key.D1:
                        tabControlMain.SelectedIndex = 0;
                        break;

                    case Key.D2:
                        tabControlMain.SelectedIndex = 1;
                        break;

                    case Key.D3:
                        tabControlMain.SelectedIndex = 2;
                        break;

                    case Key.D4:
                        tabControlMain.SelectedIndex = 3;
                        break;

                    case Key.D5:
                        tabControlMain.SelectedIndex = 4;
                        break;

                    case Key.D6:
                        tabControlMain.SelectedIndex = 5;
                        break;

                    case Key.D7:
                        tabControlMain.SelectedIndex = 6;
                        break;

                    case Key.D8:
                        tabControlMain.SelectedIndex = 7;
                        break;

                    // Carica profilo
                    case Key.O:
                        caricaConfigurazioneDalDatabaseToolStripMenuItem_Click(sender, e);
                        break;

                    // Apri log
                    case Key.L:
                        logToolStripMenu_Click(sender, e);
                        break;

                    // DB Manager
                    case Key.D:
                        gestisciDatabaseToolStripMenuItem_Click(sender, e);
                        break;

                    // Info
                    case Key.I:
                        infoToolStripMenuItem1_Click(sender, e);
                        break;

                    // Salva
                    case Key.S:

                        // Salva su database
                        if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
                        {
                            salvaConfigurazioneNelDatabaseToolStripMenuItem_Click(sender, e);
                        }
                        // Salva config. corrente
                        else
                        {
                            salvaToolStripMenuItem_Click(sender, e);
                        }

                        break;

                    // Mode
                    case Key.M:
                        if ((bool)radioButtonModeTcp.IsEnabled && (bool)radioButtonModeSerial.IsEnabled)
                        {
                            if ((bool)radioButtonModeSerial.IsChecked)
                            {
                                radioButtonModeTcp.IsChecked = true;
                            }
                            else
                            {
                                radioButtonModeSerial.IsChecked = true;
                            }
                        }
                        break;

                    // Mode
                    case Key.N:
                    case Key.B:
                        if ((bool)radioButtonModeSerial.IsChecked)
                        {
                            buttonSerialActive_Click(sender, e);
                        }
                        else
                        {
                            buttonTcpActive_Click(sender, e);
                        }
                        break;

                    // Chiudi finestra
                    case Key.W:
                    case Key.Q:
                        this.Close();
                        break;

                }

                if (e.Key == Key.C && (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift)))
                {
                    if (!statoConsole)
                    {
                        apriConsole();

                        this.Focus();
                    }
                    else
                    {
                        chiudiConsole();
                    }
                }
            }



            // Non vincolato al ctrl
            switch (e.Key)
            {
                // Cancella tabella
                case Key.Delete:

                    // Home
                    if (tabControlMain.SelectedIndex == 0)
                    {
                        buttonClearSerialStatus_Click(sender, e);
                    }

                    // Log
                    if (tabControlMain.SelectedIndex == 5)
                    {
                        buttonClearAll_Click(sender, e);
                    }

                    break;
            }
        }

        private void logToolStripMenu_Click(object sender, RoutedEventArgs e)
        {
            if (!logWindowIsOpen)
            {
                logWindowIsOpen = true;
                LogView window = new LogView(this);
                window.Show();
            }
            else
            {
                MessageBox.Show(lang.languageTemplate["strings"]["logIsAlreadyOpen"], "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void CheckBoxSwapCrcBytes_Checked(object sender, RoutedEventArgs e)
        {
            ModBus.swapCrcBytes = (bool)CheckBoxSwapCrcBytes.IsChecked;
        }
    }



    // Classe per caricare dati dal file di configurazione json
    public class SAVE
    {
        // Variabili interfaccia 
        public bool usingSerial { get; set; } //True -> Serial, False -> TCP

        public string modbusAddress { get; set; }

        //Vatiabili configurazione seriale
        public int serialPort { get; set; }
        public int serialSpeed { get; set; }
        public int serialParity { get; set; }
        public int serialStop { get; set; }

        public bool serialMaster { get; set; } //True -> Master, False -> False
        public bool serialRTU { get; set; }

        //Variabili configurazione tcp
        public string tcpClientIpAddress { get; set; }
        public string tcpClientPort { get; set; }
        public string tcpServerIpAddress { get; set; }
        public string tcpServerPort { get; set; }

        //VARIABILI TABPAGE

        //TabPage1 (Coils)
        public string textBoxCoilsAddress01 { get; set; }
        public string textBoxCoilNumber { get; set; }
        public string textBoxCoilsRange_A { get; set; }
        public string textBoxCoilsRange_B { get; set; }
        public string textBoxCoilsAddress05 { get; set; }
        public string textBoxCoilsValue05 { get; set; }
        public string textBoxCoilsAddress15_A { get; set; }
        public string textBoxCoilsAddress15_B { get; set; }
        public string textBoxCoilsValue15 { get; set; }
        public string textBoxGoToCoilAddress { get; set; }


        //TabPage2 (inputs)
        public string textBoxInputAddress02 { get; set; }
        public string textBoxInputNumber { get; set; }
        public string textBoxInputRange_A { get; set; }
        public string textBoxInputRange_B { get; set; }
        public string textBoxGoToInputAddress { get; set; }

        //TabPage3 (input registers)
        public string textBoxInputRegisterAddress04 { get; set; }
        public string textBoxInputRegisterNumber { get; set; }
        public string textBoxInputRegisterRange_A { get; set; }
        public string textBoxInputRegisterRange_B { get; set; }
        public string textBoxGoToInputRegisterAddress { get; set; }

        //TabPage4 (holding registers)
        public string textBoxHoldingAddress03 { get; set; }
        public string textBoxHoldingRegisterNumber { get; set; }
        public string textBoxHoldingRange_A { get; set; }
        public string textBoxHoldingRange_B { get; set; }
        public string textBoxHoldingAddress06 { get; set; }
        public string textBoxHoldingValue06 { get; set; }
        public string textBoxHoldingAddress16_A { get; set; }
        public string textBoxHoldingAddress16_B { get; set; }
        public string textBoxHoldingValue16 { get; set; }
        public string textBoxGoToHoldingAddress { get; set; }

        //TabPage5 (diagnostic)

        //TabPage6 (summary)
        public bool statoConsole { get; set; }

        //Altri elementi aggiunti dopo
        public object comboBoxCoilsAddress01_ { get; set; }
        public object comboBoxCoilsRange_A_ { get; set; }
        public object comboBoxCoilsRange_B_ { get; set; }
        public object comboBoxCoilsAddress05_ { get; set; }
        public object comboBoxCoilsValue05_ { get; set; }
        public object comboBoxCoilsAddress15_A_ { get; set; }
        public object comboBoxCoilsAddress15_B_ { get; set; }
        public object comboBoxCoilsValue15_ { get; set; }
        public object comboBoxInputAddress02_ { get; set; }
        public object comboBoxInputRange_A_ { get; set; }
        public object comboBoxInputRange_B_ { get; set; }
        public object comboBoxInputRegisterAddress04_ { get; set; }
        public object comboBoxInputRegisterRange_A_ { get; set; }
        public object comboBoxInputRegisterRange_B_ { get; set; }
        public object comboBoxHoldingAddress03_ { get; set; }
        public object comboBoxHoldingRange_A_ { get; set; }
        public object comboBoxHoldingRange_B_ { get; set; }
        public object comboBoxHoldingAddress06_ { get; set; }
        public object comboBoxHoldingValue06_ { get; set; }
        public object comboBoxHoldingAddress16_A_ { get; set; }
        public object comboBoxHoldingAddress16_B_ { get; set; }
        public object comboBoxHoldingValue16_ { get; set; }

        public object comboBoxCoilsRegistri_ { get; set; }
        public object comboBoxInputRegistri_ { get; set; }
        public object comboBoxInputRegRegistri_ { get; set; }
        public object comboBoxInputRegValori_ { get; set; }
        public object comboBoxHoldingRegistri_ { get; set; }
        public object comboBoxHoldingValori_ { get; set; }

        public object comboBoxCoilsAddress05_b_ { get; set; }
        public object comboBoxCoilsValue05_b_ { get; set; }
        public object comboBoxHoldingAddress06_b_ { get; set; }
        public object comboBoxHoldingValue06_b_ { get; set; }

        public object comboBoxInputOffset_ { get; set; }
        public object comboBoxInputRegOffset_ { get; set; }
        public object comboBoxHoldingOffset_ { get; set; }

        public string textBoxCoilsAddress05_b_ { get; set; }
        public string textBoxCoilsValue05_b_ { get; set; }
        public string textBoxHoldingAddress06_b_ { get; set; }
        public string textBoxHoldingValue06_b_ { get; set; }

        public string textBoxInputOffset_ { get; set; }
        public string textBoxInputRegOffset_ { get; set; }
        public string textBoxHoldingOffset_ { get; set; }

        public bool checkBoxUseOffsetInTables_ { get; set; }
        public bool checkBoxUseOffsetInTextBox_ { get; set; }
        public bool checkBoxFollowModbusProtocol_ { get; set; }
        public bool checkBoxCreateTableAtBoot_ { get; set; }
        public bool checkBoxSavePackets_ { get; set; }
        public bool checkBoxCloseConsolAfterBoot_ { get; set; }
        public bool checkBoxCellColorMode_ { get; set; }
        public bool checkBoxViewTableWithoutOffset_ { get; set; }

        public string textBoxSaveLogPath_ { get; set; }

        public object comboBoxDiagnosticFunction_ { get; set; }

        public string textBoxDiagnosticFunctionManual_ { get; set; }

        //Valori per MosBus Server
        public object comboBoxCoilsRegisters_ { get; set; }
        public object comboBoxCoilsOffset_ { get; set; }
        public string textBoxCoilsOffset_ { get; set; }

        public object comboBoxInputsRegisters_ { get; set; }
        public object comboBoxInputsOffset_ { get; set; }
        public string textBoxInputsOffset_ { get; set; }

        public object comboBoxInputsRegRegisters_ { get; set; }
        public object comboBoxInputsRegValues_ { get; set; }
        public object comboBoxInputsRegOffset_ { get; set; }
        public string textBoxInputsRegOffset_ { get; set; }

        public object comboBoxHoldingsRegisters_ { get; set; }
        public object comboBoxHoldingsValues_ { get; set; }
        public object comboBoxHoldingsOffset_ { get; set; }
        public string textBoxHoldingsOffset_ { get; set; }

        public bool checkBoxCheckModbusAddress_ { get; set; }

        public string colorDefaultReadCell_ { get; set; }
        public string colorDefaultWriteCell_ { get; set; }
        public string colorErrorCell_ { get; set; }

        public string nome_applicazione { get; set; }
        public bool checkBoxCorreggiRegistriModbus_ { get; set; }

        //tabelle
        public string[] coils_A { get; set; }
        public string[] coils_B { get; set; }
        public string[] coils_C { get; set; }

        public string[] inputs_A { get; set; }
        public string[] inputs_B { get; set; }
        public string[] inputs_C { get; set; }

        public string[] inputRegisters_A { get; set; }
        public string[] inputRegisters_B { get; set; }
        public string[] inputRegisters_C { get; set; }

        public string[] holdings_A { get; set; }
        public string[] holdings_B { get; set; }
        public string[] holdings_C { get; set; }

        public bool checkBoxColorMenu_ { get; set; }
        public string colorMenuStrip_ { get; set; }
        public string pathToConfiguration_ { get; set; }

        public bool? checkBoxPinWIndow_ { get; set; }
        public bool? checkBoxFocusReadRows_ { get; set; }
        public bool? checkBoxFocusWriteRows_ { get; set; }
        public bool? checkBoxDisableGraphics_ { get; set; }

        public string textBoxOffsetFocusTabelle_ { get; set; }
    }

    public class dataGridJson
    {
        public string[] registers { get; set; }
        public string[] values { get; set; }
        public string[] note { get; set; }
    }
}



