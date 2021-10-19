
using System;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Sockets;
using System.IO.Ports;
using System.Threading;
using System.Windows.Media;
using System.Collections.ObjectModel;
using System.Collections.Concurrent;

using ModBus_Server;

namespace ModBusMaster_Chicco
{
    public class ModBus_Chicco
    {
        //RTU o ASCII
        SerialPort serialPort;

        //TCP
        String ip_address;
        String port;

        String type;    //"RTU", "ASCII", "TCP"

        String prefix = ""; // in futuro si può aggiungere 0x o un altro prefisso al valori scritti via modbus nelle tabella holding registers

        MainWindow main;

        Border pictureBoxSending = new Border();
        Border pictureBoxReceiving = new Border();

        TcpListener server; // L'oggetto server per Modbus TCP deve essere globale

        ComboBox[] comboBoxRegisters = new ComboBox[4];
        ComboBox[] comboBoxValues = new ComboBox[4];
        ComboBox[] comboBoxOffset = new ComboBox[4];
        TextBox[] textBoxOffset = new TextBox[4];
        DataGrid[] dataGrid = new DataGrid[4];

        String[] comboBoxRegisters_str = new String[4];
        String[] comboBoxValues_str = new String[4];
        String[] comboBoxOffset_str = new String[4];
        String[] textBoxOffset_str = new String[4];
        DataGrid[] dataGrid_ = new DataGrid[4];

        SolidColorBrush[] colorDefaultCell = new SolidColorBrush[3];

        ObservableCollection<ModBus_Item>[] list;

        String ModBusAddress = "";
        public bool checkModBusAddress = true;
        public bool correctModbusRegister = true;

        //int lastHoldingRange = 0;
        //int lastHoldingAddress = 0;

        // Variabili esecuzione programma
        public bool disableGraphics = false;    // Disabilita grafica e minimizza i dispatcher
        public bool scrollToReadRows = true;
        public bool scrollToWriteRows = true;

        public int mode = 0;    // 0 -> Started, 1 -> Paused, 2 -> Stopped

        public int offsetFocusTabelle = 5;

        public FixedSizedQueue<String> log = new FixedSizedQueue<string>();
        public FixedSizedQueue<String> log2 = new FixedSizedQueue<string>();
        public FixedSizedQueue<String> logStatus = new FixedSizedQueue<string>();

        //------------------------------------------------------------------------------------------
        //-----------------------------Definizione funnzione ModBus 1-------------------------------
        //------------------------------------------------------------------------------------------

        // eliminata

        //------------------------------------------------------------------------------------------
        //-----------------------------Definizione funnzione ModBus 2-------------------------------
        //------------------------------------------------------------------------------------------

        public ModBus_Chicco(MainWindow main_, String type_)
        {
            main = main_;

            // Type: TCP, RTU, ASCII
            type = type_;

            // RTU/ASCII
            serialPort = main.serialPort;

            // TCP
            ip_address = main.textBoxTcpClientIpAddress.Text;
            port = main.textBoxTcpClientPort.Text;

            ModBusAddress = main.textBoxModbusAddress.Text;

            checkModBusAddress = (bool)main.checkBoxCheckModbusAddress.IsChecked;
            correctModbusRegister = (bool)main.checkBoxCorreggiRegistriModbus.IsChecked;

            // GRAFICA
            pictureBoxSending = main.pictureBoxIsResponding;
            pictureBoxReceiving = main.pictureBoxIsSending;

            // Dimensione log locale
            log.Limit = 10000;
            log2.Limit = 10000;
            logStatus.Limit = 10000;

            // Tabelle registri letti via ModBus
            list = new ObservableCollection<ModBus_Item>[] { main.list_coilsTable, main.list_inputsTable, main.list_holdingRegistersTable, main.list_inputRegistersTable }; ;
            comboBoxRegisters = new ComboBox[] { main.comboBoxCoilsRegisters, main.comboBoxInputsRegisters, main.comboBoxHoldingsRegisters, main.comboBoxInputsRegRegisters }; ;
            comboBoxValues = new ComboBox[] { main.comboBoxHoldingsValues, main.comboBoxInputsRegValues }; ;
            comboBoxOffset = new ComboBox[] { main.comboBoxCoilsOffset, main.comboBoxInputsOffset, main.comboBoxHoldingsOffset, main.comboBoxInputsRegOffset }; ;
            textBoxOffset = new TextBox[] { main.textBoxCoilsOffset, main.textBoxInputsOffset, main.textBoxHoldingsOffset, main.textBoxInputsRegOffset }; ;

            colorDefaultCell = new SolidColorBrush[] { main.colorDefaultWriteCell, main.colorDefaultReadCell, main.colorErrorCell }; ;

            dataGrid = new DataGrid[] { main.dataGridViewCoils, main.dataGridViewInput, main.dataGridViewHolding, main.dataGridViewInputRegister }; ;

            disableGraphics = (bool)main.checkBoxDisableGraphics.IsChecked;
            scrollToReadRows = (bool)main.checkBoxFocusReadRows.IsChecked;
            scrollToWriteRows = (bool)main.checkBoxFocusWriteRows.IsChecked;


            if (!int.TryParse(main.textBoxOffsetFocusTabelle.Text, out offsetFocusTabelle))
            {
                offsetFocusTabelle = 5;
            }

            comboBoxRegisters[0].Dispatcher.Invoke((Action)delegate
            {
                for (int i = 0; i < 4; i++)
                {
                    comboBoxRegisters_str[i] = comboBoxRegisters[i].SelectedItem.ToString().Split(' ')[1];
                    comboBoxOffset_str[i] = comboBoxOffset[i].SelectedItem.ToString().Split(' ')[1];
                    textBoxOffset_str[i] = textBoxOffset[i].ToString();

                    if (i < 2)
                    {
                        comboBoxValues_str[i] = comboBoxValues[i].SelectedItem.ToString().Split(' ')[1];
                    }
                }
            });

            //DEBUG
            Console.WriteLine("Oggeto ModBus:" + type);

            mode = 0;

            main.updateStartPauseStop();
        }

        public void stopTcpLoop()
        {
            server.Stop();
        }

        //--------------------------------------------------------------------------------------------
        //-----------------------------Funzione per Thread TCP listener-------------------------------
        //--------------------------------------------------------------------------------------------

        // Funzione da lanciare in un thread in background che si occupa dell'ascolto dei messaggi in
        // entrata e dell'invio della risposta in uscita -> TCP

        public void loopTcpListener()
        {
            server = null;

            try
            {
                // TcpListener server = new TcpListener(port);
                if (ip_address == "localhost")
                    ip_address = "0.0.0.0";

                server = new TcpListener(IPAddress.Parse(ip_address), int.Parse(port));

                // Start listening for client requests.
                server.Start();

                // Enter the listening loop.
                while (true)
                {
                    // mode != stopped
                    if (mode < 2)
                    {
                        Console.WriteLine("\nloopTcpListener:\tWaiting for a connection... ");

                        try
                        {
                            TcpClient client = server.AcceptTcpClient();

                            // Con ThreadPool
                            ThreadPool.QueueUserWorkItem(this.handleClient, client);

                            // Con Thread
                            //Thread p = new Thread(new ParameterizedThreadStart(this.handleClient));
                            //p.Priority = ThreadPriority.Highest;
                            //p.Start(client);
                        }
                        catch (Exception err)
                        {
                            // DEBUG
                            Console.WriteLine(err);
                        }
                    }
                    else
                    {
                        Thread.Sleep(100);
                    }
                }
            }
            catch (SocketException e)
            {
                Console.WriteLine("loopTcpListener:\tSocketException: {0}", e);
            }

            Console.WriteLine("loopTcpListener:\tLoop arrested");
        }

        private void handleClient(object token)
        {
            Byte[] bytes = new Byte[1024];
            StringBuilder builder = new StringBuilder();

            using (var client = token as TcpClient)
            {
                using (var stream = client.GetStream())
                {
                    Console.WriteLine("{0,20}{1}", "\nhandleClient:", "Connected!");

                    int currentTable = 0;   // 0 -> Coils, 1 -> Inputs, 2 -> Holding reg.,  4 -> Input reg.
                    int currentRow = 0;     // Riga tabella

                    Byte[] received = new Byte[256];
                    int i;

                    i = stream.Read(received, 0, received.Length);

                    if (received[7] > 0)
                    {
                        logStatus.Enqueue(DateTime.Now.ToString().Split(' ')[1] + " - " + "Incoming connection from " + IPAddress.Parse(((IPEndPoint)client.Client.RemoteEndPoint).Address.ToString()) + " - FC" + received[7].ToString().PadLeft(2, '0') + "\n");
                    }

                    receiving(received, i);

                    Console.Write("{0,20}{1}", "\nhandleClient:", "Received: ");

                    for (int a = 0; a < i; a++)
                    {
                        Console.Write(received[a].ToString() + " ");
                    }

                    stopReceiving();

                    // TCP Slave ID -> 6
                    if ((bool)checkModBusAddress && received[6] != byte.Parse(ModBusAddress))
                    {
                        MessageBox.Show("Slave ID messaggio ricevuto non valido (controllare ID del server o della richiesta del client", "Alert");
                        Console.WriteLine("Slave ID messaggio ricevuto non valido (controllare ID del server o della richiesta del client");
                    }

                    //Controllo che il pacchetto ricevuto sia approx Modbus
                    //received[0] == 0 && received[1] == 1 && received[2] == 0 && received[3] == 0
                    if (true)
                    {
                        sending();

                        byte[] response;


                        //Analizzo il pacchetto ricevuto
                        try
                        {

                            response = Modbus_Runtime(received, out currentTable, out currentRow);

                        }
                        catch (Exception error)
                        {
                            Console.WriteLine(error);
                            response = new byte[1];
                        }

                        Console.Write("{0,20}{1}", "\nhandleClient:", "Sent: ");

                        for (int a = 0; a < response.Length; a++)
                        {
                            Console.Write(response[a].ToString() + " ");
                        }

                        // Send back a response.
                        try
                        {
                            stream.Write(response, 0, response.Length);
                        }
                        catch (Exception err)
                        {
                            Console.WriteLine(err);
                        }

                        stopSending(response);
                    }

                    // Shutdown and end connection
                    client.Close();

                    // Scorro la tabella alla riga letta
                    ScrollTable(currentTable, currentRow);

                    // DoEvents();
                }
            }
            Console.WriteLine("{0,20}{1}", "\nhandleClient:", "Releasing thread handler client");
        }


        public static void DoEvents()
        {
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, new Action(delegate { }));
        }

        //--------------------------------------------------------------------------------------------
        //-----------------------------Funzione per Thread RTU listener-------------------------------
        //--------------------------------------------------------------------------------------------

        // Funzione da lanciare in un thread in background che si occupa dell'ascolto dei messaggi in
        // entrata e dell'invio della risposta in uscita -> RTU

        public void loopSerialListenerRTU()
        {
            /*
            TCP:
            0x00
            0x01
            0x00
            0x00
            0x00 -> Message Length Hi
            0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

            0x07 -> Slave Address
            0x01 -> Function
            0x01 -> Start Addr Hi
            0x2C -> Start Addr Lo
            0x00 -> No of Registers Hi
            0x03 -> No of Registers Lo
             */


            while (true)
            {
                try
                {
                    if (serialPort.BytesToRead > 0)
                    {
                        byte[] received = new byte[256];

                        int currentTable = 0;   // 0 -> Coils, 1 -> Inputs, 2 -> Holding reg.,  4 -> Input reg.
                        int currentRow = 0;     // Riga tabella

                        Thread.Sleep(50);
                        //Leggo il buffer della seriale
                        int Length = serialPort.Read(received, 0, received.Length);

                        // Aggiorno grafica received
                        receiving(received, Length);

                        Console.Write("Length: " + Length.ToString());

                        // DEBUG
                        Console_printByte("received: ", received, received.Length);

                        stopReceiving();

                        // RTU Slave ID -> 0
                        if (checkModBusAddress && received[0] != byte.Parse(ModBusAddress))
                        {
                            //MessageBox.Show("Slave ID messaggio ricevuto non valido (controllare ID del server o della richiesta del client", "Alert");
                            Console.WriteLine("Slave ID messaggio ricevuto non valido (controllare ID del server o della richiesta del client");
                        }

                        // Calcolo il CRC del messaggio ricevuto
                        byte[] CRC_received = Calcolo_CRC(received, (Length - 2));

                        // Controllo che il CRC calcolato coincida con quello contenuto nel messaggio NB: il CRC è little endian
                        if (received[Length - 1] != CRC_received[1] || received[Length - 2] != CRC_received[0])
                        {
                            //MessageBox.Show("CRC ultimo pacchetto ricevuto non valido", "Alert");
                            Console.WriteLine("CRC ultimo pacchetto ricevuto non valido");
                            Console.WriteLine("CRC pacchetto: " + received[Length - 1].ToString() + " " + received[Length - 2].ToString());
                            Console.WriteLine("CRC calcolato: " + CRC_received[1].ToString() + " " + CRC_received[0].ToString());

                            log.Enqueue("Errore CRC pacchetto, ricevuto: [" + received[Length - 1].ToString("X").PadLeft(2, '0') + " " + received[Length - 2].ToString("X").PadLeft(2, '0') + "], calcolato: [" + CRC_received[1].ToString("X").PadLeft(2, '0') + " " + CRC_received[0].ToString("X").PadLeft(2, '0') + "]");
                            log2.Enqueue("Errore CRC pacchetto, ricevuto: [" + received[Length - 1].ToString("X").PadLeft(2, '0') + " " + received[Length - 2].ToString("X").PadLeft(2, '0') + "], calcolato: [" + CRC_received[1].ToString("X").PadLeft(2, '0') + " " + CRC_received[0].ToString("X").PadLeft(2, '0') + "]");

                            //NB: Non usare return altrimenti fermo il thread di ricezione
                            //return; // Se CRC invalido non rispondo al messaggio e fermo il codice qui
                        }

                        sending();

                        Console_printByte("received: ", received, Length);


                        // Converto il dato ricevuto su RTU in ModBus TCP
                        // uso la funzione che gestisce ModBus TCP anche per seriale spostando gli elementi nell'array
                        byte[] receivedAsTcp = new byte[Length + 6];
                        Array.Copy(received, 0, receivedAsTcp, 6, Length);
                        Console_printByte("receivedAsTcp: ", receivedAsTcp, receivedAsTcp.Length);

                        byte[] responseAsTcp;

                        // Analizzo il pacchetto ricevuto
                        try
                        {
                            responseAsTcp = Modbus_Runtime(receivedAsTcp, out currentTable, out currentRow);
                        }
                        catch (Exception error)
                        {
                            responseAsTcp = new byte[6];
                            Console.WriteLine(error);
                        }

                        Console_printByte("responseAsTcp: ", responseAsTcp, responseAsTcp.Length);

                        // Converto il dato da TCP per RTU:
                        byte[] response = new byte[responseAsTcp.Length - 4];   //-6+2(CRC)

                        //DEBGU
                        Console.WriteLine("responseAsTcp.Length: " + responseAsTcp.Length.ToString());

                        Array.Copy(responseAsTcp, 6, response, 0, (responseAsTcp.Length - 6));
                        Console_printByte("response: ", response, response.Length);

                        // Calcolo il CRC del messaggio da inviare
                        byte[] CRC_send = Calcolo_CRC(response, (response.Length - 2));

                        //Aggiungo il CRC alla fine del pacchetto
                        response[(response.Length - 2)] = CRC_send[0];  // LSB CRC
                        response[(response.Length - 1)] = CRC_send[1];  // MSB CRC

                        // DEBUG
                        Console_printByte("Sent: ", response, response.Length);

                        // Send back a response.
                        serialPort.Write(response, 0, response.Length);

                        stopSending(response);

                        // Scorro la tabella alla riga letta
                        ScrollTable(currentTable, currentRow);
                    }
                }

                catch (SocketException e)
                {
                    Console.WriteLine("SocketException: {0}", e);
                }
            }
        }

        //--------------------------------------------------------------------------------------------
        //-----------------------------Funzione per Thread ASCII listener-----------------------------
        //--------------------------------------------------------------------------------------------

        // Funzione da lanciare in un thread in background che si occupa dell'ascolto dei messaggi in
        // entrata e dell'invio della risposta in uscita -> ASCII

        public void loopSerialListenerASCII()
        {
            // DA SCRIVERE
        }

        //--------------------------------------------------------------------------
        //-----------------------------MODBUS RUNTIME-------------------------------
        //--------------------------------------------------------------------------

        // Funzione che riceve il pacchetto di byte in Modbus TCP, va a leggere i valori nelle tabelle
        // e restituisce i byte di risposta
        // (per RTU basta convertire l'array letto dalla seriale nel formato TCP, passarlo alla funzione runtime
        // e convertire la risposta TCP di nuovo in RTU, aggiungere il CRC ed inviare la risposta)

        public byte[] Modbus_Runtime(byte[] received, out int currentTable, out int currentRow)
        {
            /*
            TCP:
            0x00
            0x01
            0x00
            0x00
            0x00 -> Message Length Hi
            0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

            0x07 -> Slave Address   (DA qui in poi per RTU)
            0x01 -> Function
            0x01 -> Start Addr Hi
            0x2C -> Start Addr Lo
            0x00 -> No of Registers Hi
            0x03 -> No of Registers Lo
             */

            UInt16 startAddress;            // Primo registro richiesto
            UInt16 numberOfRegisters;       // Numero di registri richiesti

            int numberOfTablesRegisters;    // Numero di righe presenti nella tabella selezionata

            UInt16[] value;                 // Array 0-9998 dove metto nella rispettiva cella di indirizzo il valore
                                            // corrispodenti nella tabella Modbus selezionata

            byte[] response;
            bool registerFound = new bool();

            currentRow = -1;                // Defualt = -1

            switch (received[7])
            {
                //-------------------------------------------------------------------
                //------------------------Read coil status---------------------------
                //-------------------------------------------------------------------
                case 1: // gestito assieme al caso 2
                //-------------------------------------------------------------------
                //------------------------Read input status--------------------------
                //-------------------------------------------------------------------
                case 2:
                    startAddress = (UInt16)(((UInt32)(received[8]) << 8) + (UInt32)(received[9]));
                    numberOfRegisters = (UInt16)(((UInt32)(received[10]) << 8) + (UInt32)(received[11]));

                    int byteCount = 0;

                    if (numberOfRegisters % 8 == 0)
                    {
                        response = new byte[numberOfRegisters / 8 + 9];
                        byteCount = numberOfRegisters / 8;
                    }
                    else
                    {
                        response = new byte[numberOfRegisters / 8 + 10];
                        byteCount = (numberOfRegisters / 8 + 1);
                    }


                    //Header
                    response[0] = received[0];
                    response[1] = received[1];
                    response[2] = received[2];  // I primi byte contengono un numero crescente di chiamate al server
                    response[3] = received[3];  // se non copiato nella risposta il PLC scarta il pacchetto

                    //Message Length
                    response[4] = (byte)((UInt16)(response.Length - 6) << 8);
                    response[5] = (byte)(response.Length - 6);

                    //Slave Address
                    response[6] = received[6];

                    //FC Code
                    response[7] = received[7];

                    //Numero di byte a seguire
                    response[8] = (byte)(byteCount);

                    //bool[] coilStatus = new bool[numberOfRegisters];
                    numberOfTablesRegisters = list[received[7] - 1].Count;

                    //Leggo i dati della tabella e li converto in array
                    value = new UInt16[65536];  // Metto 65536, per eventuali dispositivi modbus che superano il protocollo da 0 a 65535

                    for (int i = 0; i < numberOfTablesRegisters; i++)
                    {
                        // Indice modbus
                        int index = 0;

                        // DEBUG
                        //Console.WriteLine("\nPrimo registro tabella: " + (uint_parser(list[received[7] - 1][i].Register, comboBoxRegisters[received[7] - 1]) + uint_parser(textBoxOffset_[received[7] - 1], comboBoxOffset[received[7] - 1])).ToString());

                        if (!(bool)correctModbusRegister) // False -> tabelle senza offset / True -> tabelle con offset
                        {
                            index = uint_parser(list[received[7] - 1][i].Register, comboBoxRegisters_str[received[7] - 1]) + uint_parser(textBoxOffset_str[received[7] - 1], comboBoxOffset_str[received[7] - 1]);

                            value[index] = UInt16.Parse(list[received[7] - 1][i].Value);
                        }
                        else if (received[7] == 1)
                        {
                            index = uint_parser(list[received[7] - 1][i].Register, comboBoxRegisters_str[received[7] - 1]) + uint_parser(textBoxOffset_str[received[7] - 1], comboBoxOffset_str[received[7] - 1]) - 1;

                            // Coils (Offset 1)
                            value[index] = UInt16.Parse(list[received[7] - 1][i].Value);
                        }
                        else if (received[7] == 2)
                        {
                            index = uint_parser(list[received[7] - 1][i].Register, comboBoxRegisters_str[received[7] - 1]) + uint_parser(textBoxOffset_str[received[7] - 1], comboBoxOffset_str[received[7] - 1]) - 10001;
                            // Inputs (Offset 10001)
                            value[index] = UInt16.Parse(list[received[7] - 1][i].Value);
                        }

                        // Metto lo sfondo alle celle, colorato se fa parte di quelle lette o bianco se fuori
                        if (index >= startAddress && index < (startAddress + numberOfRegisters))
                        {
                            dataGrid[received[7] - 1].Dispatcher.Invoke((Action)delegate
                            {
                                list[received[7] - 1][i].Color = colorDefaultCell[1].ToString();
                            });
                        }
                        else
                        {
                            dataGrid[received[7] - 1].Dispatcher.Invoke((Action)delegate
                            {
                                list[received[7] - 1][i].Color = Brushes.White.ToString();
                            });
                        }

                        // Controllo tra tutti i registri cercati quello che trovo (vale per FC01 FC02 FC03 FC04, valutare se aggiungerlo a FC15 e FC16)
                        for (int a = 0; a < numberOfRegisters; a++)
                        {
                            if (index == (startAddress + a) && scrollToReadRows && currentRow < 0)
                            {
                                currentRow = i;
                            }
                        }
                    }

                    // Costruzione pacchetto in risposta
                    for (int a = 0; a < (numberOfRegisters / 8 + 1); a++)
                    {
                        for (int i = 0; i < 8; i++)
                        {
                            try
                            {
                                if (value[i + a * 8 + startAddress] > 0)
                                {
                                    response[9 + a] += (byte)(1 << i);
                                }
                            }
                            catch { }
                        }
                    }

                    /*comboBoxOffset[0].Dispatcher.Invoke((Action)delegate
                    {
                        lock (list)
                        {
                            dataGrid[received[7] - 1].Items.Refresh();
                        }
                    });*/

                    currentTable = received[7] - 1;

                    return response;



                //-------------------------------------------------------------------
                //----------------------Read holding register FC04-------------------
                //-------------------------------------------------------------------
                case 4: // gestito assieme al caso 3
                //-------------------------------------------------------------------
                //-----------------------Read input register FC03--------------------
                //-------------------------------------------------------------------
                case 3:

                    numberOfTablesRegisters = 0;

                    startAddress = (UInt16)(((UInt32)(received[8]) << 8) + (UInt32)(received[9]));
                    numberOfRegisters = (UInt16)(((UInt32)(received[10]) << 8) + (UInt32)(received[11]));

                    response = new byte[numberOfRegisters * 2 + 9];

                    // Header
                    response[0] = received[0];
                    response[1] = received[1];
                    response[2] = received[2];
                    response[3] = received[3];

                    // Message Length
                    response[4] = (byte)((UInt16)(response.Length - 6) << 8);
                    response[5] = (byte)(response.Length - 6);

                    // Slave Address
                    response[6] = received[6];

                    // FC Code
                    response[7] = received[7];

                    // Byte Count
                    response[8] = (byte)(numberOfRegisters * 2);  // Ogni registro richiede 2 byte

                    UInt16[] regStatus = new UInt16[numberOfRegisters];

                    // DEBUG, prova per velocizzare lettura registri
                    //--------------------------------------------------------------------------------------

                    //Leggo i dati della tabella e li converto in un array con i valori di indice corretto
                    value = new UInt16[65536];  // Metto 65536, per eventuali dispositivi modbus che superano il protocollo da 0 a 65535

                    if (!disableGraphics)
                    {
                        dataGrid[0].Dispatcher.Invoke((Action)delegate
                        {
                            for (int i = 0; i < list[received[7] - 1].Count; i++)
                            {
                                //list[received[7] - 1][i].Color = null;
                                list[received[7] - 1][i].Color = Brushes.White.ToString();
                            }
                        });
                    }

                    for (int i = 0; i < list[received[7] - 1].Count; i++)
                    {
                        // Indice modbus
                        int index = 0;

                        if (!(bool)correctModbusRegister)
                        {
                            index = uint_parser(list[received[7] - 1][i].Register, comboBoxRegisters_str[received[7] - 1]) + uint_parser(textBoxOffset_str[received[7] - 1], comboBoxOffset_str[received[7] - 1]);

                            value[index]
                                = uint_parser(list[received[7] - 1][i].Value, comboBoxValues_str[received[7] - 3]);
                        }
                        else if (received[7] == 4)
                        {
                            index = uint_parser(list[received[7] - 1][i].Register, comboBoxRegisters_str[received[7] - 1]) + uint_parser(textBoxOffset_str[received[7] - 1], comboBoxOffset_str[received[7] - 1]) - 30001;

                            value[index]
                                = uint_parser(list[received[7] - 1][i].Value, comboBoxValues_str[received[7] - 3]);
                        }
                        else if (received[7] == 3)
                        {
                            index = uint_parser(list[received[7] - 1][i].Register, comboBoxRegisters_str[received[7] - 1]) + uint_parser(textBoxOffset_str[received[7] - 1], comboBoxOffset_str[received[7] - 1]) - 40001;

                            value[index]
                                = uint_parser(list[received[7] - 1][i].Value, comboBoxValues_str[received[7] - 3]);
                        }

                        // Metto lo sfondo alle celle, colorato se fa parte di quelle lette o bianco se fuori
                        if (index >= startAddress && index < (startAddress + numberOfRegisters))
                        {
                            dataGrid[received[7] - 1].Dispatcher.Invoke((Action)delegate
                            {
                                list[received[7] - 1][i].Color = colorDefaultCell[1].ToString();
                            });
                        }
                        else
                        {
                            /*dataGrid[received[7] - 1].Dispatcher.Invoke((Action)delegate
                            {
                                list[received[7] - 1][i].Color = Brushes.White.ToString();
                            });*/
                        }

                        /*if(index >= lastHoldingAddress && index < (lastHoldingAddress + lastHoldingRange))
                        {
                            dataGrid[received[7] - 1].Dispatcher.Invoke((Action)delegate
                            {
                                list[received[7] - 1][i].Color = null;// Brushes.White.ToString();
                            });
                        }*/

                        // Controllo tra tutti i registri cercati quello che trovo (se cercavo solo startaddress (primo registro e non lo trovava non scrollava la tabella))
                        for (int a = 0; a < numberOfRegisters; a++)
                        {
                            if (index == (startAddress + a) && scrollToReadRows && currentRow < 0)
                            {
                                currentRow = i;
                            }
                        }
                    }

                    /*dataGrid[received[7] - 1].Dispatcher.Invoke((Action)delegate
                    {
                        dataGrid[received[7] - 1].Items.Refresh();
                    });*/

                    //lastHoldingAddress = startAddress;
                    //lastHoldingRange = numberOfRegisters;

                    // debug
                    //Console.WriteLine("lastHoldingAddress: " + lastHoldingAddress.ToString());
                    //Console.WriteLine("lastHoldingRange: " + lastHoldingRange.ToString());


                    for (int a = 0; a < numberOfRegisters; a += 1)
                    {
                        response[8 + a * 2 + 1] += (byte)(value[a + startAddress] >> 8);
                        response[8 + a * 2 + 2] += (byte)(value[a + startAddress]);
                    }

                    /*dataGrid[received[7] - 1].Dispatcher.Invoke((Action)delegate
                    {
                        dataGrid[received[7] - 1].ItemsSource = null;
                        dataGrid[received[7] - 1].ItemsSource = list[received[7] - 1];
                    });*/

                    /*if (!disableGraphics)
                    {
                        dataGrid[received[7] - 1].Dispatcher.Invoke((Action)delegate
                        {
                            dataGrid[2].Items.Refresh();
                            //dataGrid[received[7] - 1].CommitEdit();
                        });
                    }*/

                    currentTable = received[7] - 1;

                    return response;


                //-------------------------------------------------------------------
                //------------------------Force single coil--------------------------
                //-------------------------------------------------------------------
                case 5:
                    startAddress = (UInt16)(((UInt32)(received[8]) << 8) + (UInt32)(received[9]));
                    numberOfRegisters = (UInt16)(((UInt32)(received[10]) << 8) + (UInt32)(received[11]));

                    registerFound = false;

                    for (int i = 0; i < list[0].Count; i++)    //Passo fuori le righe
                    {
                        if ((uint_parser(list[0][i].Register, comboBoxRegisters_str[0]) + uint_parser(textBoxOffset_str[0], comboBoxOffset_str[0])) == (UInt16)(startAddress))
                        {
                            Console.Write("\nTrovata corrispondenza: " + (startAddress).ToString());

                            comboBoxOffset[0].Dispatcher.Invoke((Action)delegate
                            {
                                lock (list)
                                {
                                    if (numberOfRegisters == 0xFF00)
                                    {
                                        list[0][i].Value = "1";
                                        list[0][i].Color = colorDefaultCell[0].ToString();
                                    }
                                    else
                                    {
                                        list[0][i].Value = "0";
                                        list[0][i].Color = colorDefaultCell[0].ToString();
                                    }
                                }
                            });

                            if (scrollToWriteRows)
                            {
                                currentRow = i;
                            }

                            registerFound = true;
                        }
                        else
                        {

                            //list[0][i].Color = Brushes.White.ToString();
                        }
                    }

                    // Se alla fine del for non ho trovata corrispondenza, aggiungo il registro in fondo alla tabella
                    if (!registerFound)
                    {
                        Console.WriteLine("Nessuna corrispondenza trovata, aggiungo nuova riga: " + (startAddress).ToString());

                        ModBus_Item row = new ModBus_Item();


                        row.Register = (startAddress - uint_parser(textBoxOffset_str[0], comboBoxOffset_str[0])).ToString();

                        comboBoxOffset[0].Dispatcher.Invoke((Action)delegate
                        {
                            if (numberOfRegisters == 0xFF00)
                            {
                                row.Value = "1";
                            }
                            else
                            {
                                row.Value = "0";
                            }

                            row.Color = colorDefaultCell[0].ToString();

                            lock (list)
                            {
                                list[0].Add(row);
                            }

                            //dataGrid[0].Items.Refresh();
                        });

                        if (scrollToWriteRows)
                        {
                            currentRow = list[0].Count() - 1;
                        }
                    }

                    // Aggiorno la raccolta di items nella tabella
                    /*if (!disableGraphics)
                    {
                        dataGrid[0].Dispatcher.Invoke((Action)delegate
                        {
                            dataGrid[0].Items.Refresh();
                        });
                    }*/


                    // La risposta è una semplice echo della query, non posso usare direttamente
                    // received perche' e' lunga 256 caratteri
                    response = new byte[12];

                    Array.Copy(received, 0, response, 0, 12);

                    currentTable = 0;

                    return response;


                //-------------------------------------------------------------------
                //---------------------Preset holding register-----------------------
                //-------------------------------------------------------------------
                case 6:
                    startAddress = (UInt16)(((UInt32)(received[8]) << 8) + (UInt32)(received[9]));
                    numberOfRegisters = (UInt16)(((UInt32)(received[10]) << 8) + (UInt32)(received[11]));

                    registerFound = false;

                    for (int i = 0; i < list[2].Count; i++)    //Passo fuori le righe
                    {
                        if (((uint_parser(list[2][i].Register, comboBoxRegisters_str[2]) + uint_parser(textBoxOffset_str[2], comboBoxOffset_str[2])) == (UInt16)(startAddress) && !(bool)correctModbusRegister) ||
                            ((uint_parser(list[2][i].Register, comboBoxRegisters_str[2]) + uint_parser(textBoxOffset_str[2], comboBoxOffset_str[2])) == ((UInt16)(startAddress) + 40001) && (bool)correctModbusRegister))
                        {
                            Console.WriteLine("Trovata corrispondenza: " + (startAddress).ToString());


                            dataGrid[0].Dispatcher.Invoke((Action)delegate
                            {
                                lock (list)
                                {
                                    if (comboBoxValues_str[0] == "DEC")
                                        list[2][i].Value = numberOfRegisters.ToString();
                                    else
                                        list[2][i].Value = prefix + numberOfRegisters.ToString("X").PadLeft(4, '0');

                                    list[2][i].Color = colorDefaultCell[0].ToString();
                                }
                            });

                            if (scrollToWriteRows)
                            {
                                currentRow = i;
                            }

                            registerFound = true;
                        }
                        else
                        {
                            /*dataGrid[0].Dispatcher.Invoke((Action)delegate
                            {
                                list[2][i].Color = Brushes.White.ToString();
                            });*/
                        }
                    }

                    // Se alla fine del for non ho trovata corrispondenza, aggiungo il registro in fondo alla tabella
                    if (!registerFound)
                    {
                        Console.WriteLine("Nessuna corrispondenza trovata, aggiungo nuova riga: " + (startAddress).ToString());

                        ModBus_Item row = new ModBus_Item();

                        if (!(bool)correctModbusRegister)
                        {
                            //Tabella senza visualizzazione registri con offset modbus (0 al posto di 40001)
                            row.Register = (startAddress - uint_parser(textBoxOffset_str[0], comboBoxOffset_str[0])).ToString();
                        }
                        else
                        {
                            //Tabella con visualizzazione registri con offset modbus (40001 al posto di 0)
                            row.Register = (startAddress - uint_parser(textBoxOffset_str[0], comboBoxOffset_str[0]) + 40001).ToString();
                        }

                        /*
                        if (comboBoxRegisters[2].SelectedItem.ToString() == "DEC")
                            row.Value = startAddress.ToString();
                        else
                            row.Value = prefix + (startAddress - uint_parser(textBoxOffset_str[2], comboBoxOffset[2])).ToString("X").PadLeft(4, '0');
                            */

                        if (comboBoxValues_str[0] == "DEC")
                            row.Value = numberOfRegisters.ToString();
                        else
                            row.Value = prefix + numberOfRegisters.ToString("X").PadLeft(4, '0');

                        // Aggiorno la raccolta di items nella tabella
                        comboBoxOffset[0].Dispatcher.Invoke((Action)delegate
                        {
                            lock (list)
                            {
                                row.Color = colorDefaultCell[0].ToString();
                                list[2].Add(row);
                            }
                        });

                        if (scrollToWriteRows)
                        {
                            currentRow = list[2].Count() - 1;
                        }
                    }

                    /*if (!disableGraphics)
                    {
                        dataGrid[2].Dispatcher.Invoke((Action)delegate
                        {
                            dataGrid[2].Items.Refresh();
                            //dataGrid[2].CommitEdit();
                        });
                    }*/

                    // La risposta è una semplice echo della query, non posso usare direttamente
                    // received perche' e' lunga 256 caratteri
                    response = new byte[12];

                    Array.Copy(received, 0, response, 0, 12);

                    currentTable = 2;

                    return response;

                //-------------------------------------------------------------------
                //------------------------Force multiple coil------------------------
                //-------------------------------------------------------------------
                case 15:
                    response = new byte[12];
                    //Header
                    response[0] = received[0];
                    response[1] = received[1];
                    response[2] = received[2];  // I primi byte contengono un numero crescente di chiamate al server
                    response[3] = received[3];  // se non copiato nella risposta il PLC scarta il pacchetto

                    //Message Length
                    response[4] = (byte)((UInt16)(response.Length - 6) << 8);
                    response[5] = (byte)(response.Length - 6);

                    //Slave Address
                    response[6] = received[6];

                    //FC Code
                    response[7] = received[7];

                    response[8] = received[8];
                    response[9] = received[9];

                    response[10] = received[10];
                    response[11] = received[11];    // La risposta è una echo della query

                    // Indirizzo iniziale
                    startAddress = (UInt16)(((UInt32)(received[8]) << 8) + (UInt32)(received[9]));

                    // Numero di coils da settare
                    numberOfRegisters = (UInt16)(((UInt32)(received[10]) << 8) + (UInt32)(received[11]));

                    // Metto lo sfondo bianco a tutte le righe
                    dataGrid[0].Dispatcher.Invoke((Action)delegate
                    {
                        lock (list)
                        {
                            for (int i = 0; i < (list[0].Count); i++)    //Passo fuori le righe
                            {
                                list[0][i].Color = Brushes.White.ToString();
                            }
                        }
                    });

                    // Passo fuori le righe
                    for (int a = 0; a < numberOfRegisters; a++)
                    {
                        registerFound = false;

                        for (int i = 0; i < (list[0].Count); i++)    //Passo fuori le righe
                        {
                            // Controllo se è presente il registro attuale nella tabella
                            if ((uint_parser(list[0][i].Register, comboBoxRegisters_str[0]) + uint_parser(textBoxOffset_str[0], comboBoxOffset_str[0])) == ((UInt16)(startAddress) + a))
                            {
                                Console.WriteLine("Trovata corrispondenza: " + (startAddress + i).ToString());

                                dataGrid[0].Dispatcher.Invoke((Action)delegate
                                {
                                    lock (list)
                                    {
                                        if ((received[13 + a / 8] & (1 << (a % 8))) > 0)
                                        {
                                            list[0][i].Value = "1";
                                            list[0][i].Color = colorDefaultCell[0].ToString();
                                        }
                                        else
                                        {
                                            list[0][i].Value = "0";
                                            list[0][i].Color = colorDefaultCell[0].ToString();
                                        }
                                    }
                                });

                                if (scrollToWriteRows && currentRow == -1)
                                {
                                    currentRow = i;
                                }

                                registerFound = true;
                            }
                        }

                        // Se alla fine non ho trovato il registro nella tabella lo aggiungo in fondo
                        if (!registerFound)
                        {
                            Console.WriteLine("Nessuna corrispondenza trovata, aggiungo nuova riga: " + (startAddress).ToString());

                            ModBus_Item row = new ModBus_Item();

                            row.Register = (startAddress - uint_parser(textBoxOffset_str[0], comboBoxOffset_str[0]) + a).ToString();
                            
                            dataGrid[0].Dispatcher.Invoke((Action)delegate
                            {
                                lock (list)
                                {
                                    if ((received[13 + a / 8] & (1 << (a % 8))) > 0)
                                    {
                                        row.Value = "1";
                                        row.Color = colorDefaultCell[0].ToString();
                                    }
                                    else
                                    {
                                        row.Value = "0";
                                        row.Color = colorDefaultCell[0].ToString();
                                    }

                            
                                    list[0].Add(row);
                                }
                            });

                            if (scrollToWriteRows && currentRow == -1)
                            {
                                currentRow = list[0].Count() - 1;
                            }
                        }
                    }

                    currentTable = 0;

                    return response;

                //-------------------------------------------------------------------
                //--------------------Preset multiple registers----------------------
                //-------------------------------------------------------------------
                case 16:
                    response = new byte[12];
                    //Header
                    response[0] = received[0];
                    response[1] = received[1];
                    response[2] = received[2];  // I primi byte contengono un numero crescente di chiamate al server
                    response[3] = received[3];  // se non copiato nella risposta il PLC scarta il pacchetto

                    //Message Length
                    response[4] = (byte)((UInt16)(response.Length - 6) << 8);
                    response[5] = (byte)(response.Length - 6);

                    //Slave Address
                    response[6] = received[6];

                    //FC Code
                    response[7] = received[7];

                    response[8] = received[8];      // Starting address Hi
                    response[9] = received[9];      // Starting address Lo
                    response[10] = received[10];    // Number of register Hi
                    response[11] = received[11];    // Number of register Lo

                    // Indirizzo iniziale
                    startAddress = (UInt16)(((UInt32)(received[8]) << 8) + (UInt32)(received[9]));

                    // Numero di registri da settare
                    numberOfRegisters = (UInt16)(((UInt32)(received[10]) << 8) + (UInt32)(received[11]));

                    // Metto lo sfondo ai registri
                    dataGrid[2].Dispatcher.Invoke((Action)delegate
                    {
                        lock (list)
                        {
                            for (int i = 0; i < (list[2].Count); i++)
                            {
                                list[2][i].Color = Brushes.White.ToString();
                            }
                        }
                    });

                    // Passo fuori le righe della tabella
                    for (int a = 0; a < numberOfRegisters; a++)
                    {
                        registerFound = false;

                        for (int i = 0; i < (list[2].Count); i++)
                        {
                            // Controllo se è presente il registro attuale nella tabella
                            if (((uint_parser(list[2][i].Register, comboBoxRegisters_str[2]) + uint_parser(textBoxOffset_str[2], comboBoxOffset_str[2])) == (UInt16)(startAddress + a) && !(bool)correctModbusRegister) ||
                                ((uint_parser(list[2][i].Register, comboBoxRegisters_str[2]) + uint_parser(textBoxOffset_str[2], comboBoxOffset_str[2])) == ((UInt16)(startAddress + a) + 40001) && (bool)correctModbusRegister))
                            {
                                // debug
                                //Console.WriteLine("Trovata corrispondenza: " + (startAddress).ToString());

                                dataGrid[2].Dispatcher.Invoke((Action)delegate
                                {
                                    lock (list)
                                    {
                                        if (comboBoxValues[0].SelectedIndex == 0)
                                            list[2][i].Value = (((UInt16)(received[13 + a * 2]) << 8) + (UInt16)(received[13 + a * 2 + 1])).ToString();
                                        else
                                            list[2][i].Value = prefix + (((UInt16)(received[13 + a * 2]) << 8) + (UInt16)(received[13 + a * 2 + 1])).ToString("X").PadLeft(4, '0');

                                        list[2][i].Color = colorDefaultCell[0].ToString();
                                    }
                                });

                                if (scrollToWriteRows && currentRow == -1)
                                {
                                    currentRow = i;
                                }

                                registerFound = true;
                            }
                        }

                        // Se alla fine non ho trovato il registro nella tabella lo aggiungo in fondo
                        if (!registerFound)
                        {
                            // debug
                            Console.WriteLine("Nessuna corrispondenza trovata, aggiungo nuova riga: " + (startAddress + a).ToString());

                            ModBus_Item row = new ModBus_Item();

                            if (!(bool)correctModbusRegister)
                            {
                                //Tabella senza visualizzazione registri con offset modbus (0 al posto di 40001)
                                row.Register = (startAddress - uint_parser(textBoxOffset_str[0], comboBoxOffset_str[0]) + a).ToString();
                            }
                            else
                            {
                                //Tabella con visualizzazione registri con offset modbus (40001 al posto di 0)
                                row.Register = (startAddress - uint_parser(textBoxOffset_str[0], comboBoxOffset_str[0]) + 40001 + a).ToString();
                            }

                            // Inserisco il valore del registro nella tabella
                            dataGrid[2].Dispatcher.Invoke((Action)delegate
                            {
                                lock (list)
                                {
                                    if (comboBoxValues[1].SelectedIndex == 0)
                                        row.Value = (((UInt16)(received[13 + a * 2]) << 8) + (UInt16)(received[13 + a * 2 + 1])).ToString();
                                    else
                                        row.Value = prefix + (((UInt16)(received[13 + a * 2]) << 8) + (UInt16)(received[13 + a * 2 + 1])).ToString("X").PadLeft(4, '0');

                                    row.Color = colorDefaultCell[0].ToString();
                                }
                            });

                            dataGrid[2].Dispatcher.Invoke((Action)delegate
                            {
                                lock (list)
                                {
                                    list[2].Add(row);
                                }
                            });

                            if (scrollToWriteRows && currentRow == -1)
                            {
                                currentRow = list[2].Count() - 1;
                            }
                        }
                    }

                    currentTable = 2;

                    return response;
            }

            currentTable = -1;

            return new byte[4];
        }

        //-----------------------------------------------------------------
        //--------------------Calcolo CRC 16 MODBUS------------------------
        //-----------------------------------------------------------------

        // Calcolo CRC metodo MODBUS
        byte[] Calcolo_CRC(byte[] message, int length)
        {
            UInt16 crc = 0xFFFF;
            byte[] result = new byte[2];

            for (int pos = 0; pos < length; pos++)
            {
                crc ^= (UInt16)message[pos];    //XOR

                for (int i = 8; i != 0; i--)
                {
                    //Passo ogni byte del messaggio
                    if ((crc & 0x0001) != 0)
                    {
                        crc >>= 1;   // Shift right and XOR 0xA001
                        crc ^= 0xA001;
                    }
                    else             // Else LSB is not set
                        crc >>= 1;   // Just shift right
                }
            }

            result[0] = (byte)(crc);        //LSB
            result[1] = (byte)(crc >> 8);   //MSB

            return result;
        }

        // Funzione che restituisce stringa pronta con hh:mm:ss (8 caratteri fissi 06:05:30)
        public string timestamp()
        {
            return DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" +
                   DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" +
                   DateTime.Now.Second.ToString().PadLeft(2, '0');
        }

        // Funzione che prende in ingresso stringa e comboBox (DEC;HEX) e restituisce il
        // valore in UInt16 convertito
        public UInt16 uint_parser(String text, String comboBox)
        {
            if (comboBox == "HEX")
            {
                if (text.IndexOf("0x") != -1 || text.IndexOf("0X") != -1)
                {
                    text = text.Replace('x', '0');
                    text = text.Replace('X', '0');
                }

                //Numero passato in hex
                try
                {
                    return UInt16.Parse(text, System.Globalization.NumberStyles.HexNumber);
                }
                catch
                {
                    //MessageBox.Show("Valore inserito non valido", "Alert");
                    return 0;
                }
            }
            else
            {
                //Numero passato in decimale
                try
                {
                    return UInt16.Parse(text);
                }
                catch
                {
                    //MessageBox.Show("Valore inserito non valido", "Alert");
                    return 0;
                }
            }
        }

        //-------------------------------------------------------------------------------------
        //----------------Funzioni stampa su console o textBox array di byte-------------------
        //------------------------------------------------------------------------------------
        private void Console_printByte(String intestazione, byte[] query, int Length)
        {
            if (Length > 0)
            {
                String message = "";
                String aa = "";

                for (int i = 0; i < Length; i++)
                {

                    aa = query[i].ToString("X");

                    if (aa.Length < 2)
                        aa = "0" + aa;

                    message += "0x" + aa + " ";
                }
                Console.WriteLine(intestazione + message);
            }
        }

        public string Console_print(string header, byte[] query, int Length)
        {
            if (Length > 0)
            {
                String message = "";
                String aa = "";

                for (int i = 0; i < Length; i++)
                {

                    aa = query[i].ToString("X");

                    if (aa.Length < 2)
                        aa = "0" + aa;

                    message += "" + aa + " ";
                }

                log.Enqueue(timestamp() + header + message + "\n");
                log2.Enqueue(timestamp() + header + message + "\n");

                return timestamp() + header + message + "\n";
            }
            else
            {
                return "";
            }
        }

        //-------------------------------------------------------------------------------------
        //----------------Funzioni aggiornamento icone lettura e scrittura-------------------
        //------------------------------------------------------------------------------------
        public void sending()
        {
            if (!disableGraphics)
            {
                pictureBoxSending.Dispatcher.Invoke((Action)delegate
                {
                    //------------pictureBox gialla-------------
                    pictureBoxSending.Background = Brushes.Yellow;
                });
            }

            if (type != "TCP")
            {
                Thread.Sleep(50);
            }
            //------------------------------------------            
        }

        public void stopSending(byte[] query)
        {
            if (!disableGraphics)
            {
                pictureBoxSending.Dispatcher.Invoke((Action)delegate
                {
                    //------------pictureBox grigia-------------
                    pictureBoxSending.Background = Brushes.LightGray;
                });
            }
                
            Console_print(" Tx -> ", query, query.Length);

            if (type != "TCP")
            {
                Thread.Sleep(50);
            }
            //------------------------------------------         
        }

        public void receiving(byte[] response, int Length)
        {
            if (!disableGraphics)
            {
                pictureBoxReceiving.Dispatcher.Invoke((Action)delegate
                {
                    //------------pictureBox gialla-------------
                    pictureBoxReceiving.Background = Brushes.Yellow;
                });
            }

            Console_print(" Rx <- ", response, Length);

            if (type != "TCP")
            {
                Thread.Sleep(50);
            }
            //-----------------------------------------            
        }

        public void stopReceiving()
        {
            if (!disableGraphics)
            {
                pictureBoxReceiving.Dispatcher.Invoke((Action)delegate
                {
                    //------------pictureBox grigia-------------
                    pictureBoxReceiving.Background = Brushes.LightGray;
                });
            }

            Thread.Sleep(50);
            //------------------------------------------        
        }

        // Funzione che scorre la tabella passata alla riga richiesta
        // (usata per scorrere le tabelle nel punto in cui un mstare ModBus lo sta interrogando)
        public void ScrollTable(int currentTable, int currentRow)
        {
            if (!disableGraphics && mode == 0)
            {
                if (currentTable >= 0 && currentRow >= 0)
                {
                    dataGrid[currentTable].Dispatcher.Invoke((Action)delegate
                    {
                        // debug
                        Console.WriteLine("CurrentTable: " + currentTable.ToString());
                        Console.WriteLine("CurrentRow: " + currentRow.ToString());

                        dataGrid[currentTable].ScrollIntoView(dataGrid[currentTable].Items[dataGrid[currentTable].Items.Count - 1]);
                        dataGrid[currentTable].UpdateLayout();

                        if (currentRow > offsetFocusTabelle)
                        {
                            currentRow = currentRow - offsetFocusTabelle;
                        }
                        else
                        {
                            currentRow = 0;
                        }

                        dataGrid[currentTable].ScrollIntoView(list[currentTable][currentRow]);
                    });
                }
                else
                {
                    // debug
                    Console.WriteLine("CurrentTable: " + currentTable.ToString());
                    Console.WriteLine("CurrentRow: " + currentRow.ToString());
                }
            }
        }
    }

    public class FixedSizedQueue<T>
    {
        ConcurrentQueue<T> q = new ConcurrentQueue<T>();
        private object lockObject = new object();

        public int Limit { get; set; }
        public void Enqueue(T obj)
        {
            q.Enqueue(obj);

            lock (lockObject)
            {
                T overflow;
                while (q.Count > Limit && q.TryDequeue(out overflow)) ;
            }
        }

        public bool TryDequeue(out T obj)
        {
            if (q.TryDequeue(out obj))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }

    public class ModBus_Item
    {
        public string Register { get; set; }
        public string Value { get; set; }
        public string Notes { get; set; }
        public string Color { get; set; }
    }
}
