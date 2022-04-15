using System;
using System.IO.Ports;
using System.Windows.Forms;
using System.Threading;

namespace _7SC34Demo
{
    public partial class Form1 : Form
    {
        public bool BlnConnect;                                 //Connection Status
        public static SerialPort SCPort = null;                 //Define serial port
        public string StrReceiver;                              //Receive the string from controller
        private bool BlnBusy;                                   //If controller is busy
        public bool BlnReadCom;                                 //If reading is finished, return TRUE
        public bool BlnStopCommand;                             //Stop waiting
        public short ShrPort;                                   //The serial port number
        private bool BlnSet;                                    //If the command sent is a set command or an inquiry command. TRUE is a set command
        private double DblPulseEqui;                            //Pulse equivalent
        int sSpeed;                                           //Current speed
        long lCurrStep;                                         //Current steps
        double dCurrPosiX;                                      //Current position X
        double dCurrPosiY;                                      //Current position Y

        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SCPort = new SerialPort();
            sSpeed = 20;
            DblPulseEqui = 1;
        }

        public void ConnectPort(short sPort)
        {
            if (SCPort.IsOpen == true) SCPort.Close();
            SCPort.PortName = "COM" + sPort.ToString();            //Set the serial port number
            SCPort.BaudRate = 9600;                                //Set the bit rate
            SCPort.DataBits = 8;                                   //Set the data bits
            SCPort.StopBits = StopBits.One;                        //Set the stop bit
            SCPort.Parity = Parity.None;                           //Set the Parity
            SCPort.ReadBufferSize = 2048;
            SCPort.WriteBufferSize = 1024;
            SCPort.DtrEnable = true;
            SCPort.Handshake = Handshake.None;
            SCPort.ReceivedBytesThreshold = 1;
            SCPort.RtsEnable = false;

            //This delegate should be a trigger event for fetching data asynchronously, it will be triggered when there is data passed from serial port.
            SCPort.DataReceived += new SerialDataReceivedEventHandler(SCPort_DataReceived);     //DataReceivedEvent delegate
            try
            {
                SCPort.Open();                                     //Open serial port
                if (SCPort.IsOpen)
                {
                    StrReceiver = "";
                    BlnBusy = true;
                    BlnSet = false;
                    SendCommand("?R\r");                           //Connect to the controller
                    Delay(10000);
                    BlnBusy = false;

                    if (StrReceiver == "?R\rOK\n")
                    {
                        BlnConnect = true;                        //Connected successfully
                        ShrPort = sPort;                          //Setial port number
                        label4.Text = "Connected Successfully";
                    }
                    else
                    {
                        BlnBusy = false;
                        BlnConnect = false;
                        label4.Text = "Failed to connect";
                        MessageBox.Show("Failed to connect", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void SCPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //****************************************************************
            //Function: SCPort_DataReceived
            //Parameters: 
            //Description: receive the data sent from serial port and handle
            //Return:
            //****************************************************************
            try
            {
                string sCurString = "";
                //Loop to receive the data from serial port
                System.Threading.Thread.Sleep(200);
                sCurString = SCPort.ReadExisting();
                if (sCurString != "")
                    StrReceiver = sCurString;
                if (BlnSet == true)
                {
                    if (StrReceiver.Length == 3)
                    {
                        if (StrReceiver.Substring(StrReceiver.Length - 3) == "OK\n")
                            BlnReadCom = true;
                    }
                    else if (StrReceiver.Length == 4)
                    {
                        if (StrReceiver.Substring(StrReceiver.Length - 3) == "OK\n" || StrReceiver.Substring(StrReceiver.Length - 4) == "OK\nS")
                            BlnReadCom = true;
                    }
                    else
                    {
                        if (StrReceiver.Substring(StrReceiver.Length - 3) == "OK\n" || StrReceiver.Substring(StrReceiver.Length - 4) == "OK\nS" ||
                            StrReceiver.Substring(StrReceiver.Length - 5) == "ERR1\n" || StrReceiver.Substring(StrReceiver.Length - 5) == "ERR3\n" ||
                            StrReceiver.Substring(StrReceiver.Length - 5) == "ERR4\n" || StrReceiver.Substring(StrReceiver.Length - 5) == "ERR5\n")
                            BlnReadCom = true;
                    }
                }
                else
                {
                    if (StrReceiver.Substring(StrReceiver.Length - 1, 1) == "\n")
                        BlnReadCom = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to receive data", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public void ClosePort()
        {
            //****************************************************************
            //Function: ClosePort
            //Parameters: 
            //Description: close the connection
            //Return:
            //****************************************************************
            if (SCPort.IsOpen) SCPort.Close();
        }

        public void SendCommand(string CommandString)
        {
            //****************************************************************
            //Function: SendCommand
            //Parameters: CommandString: the command string
            //Description: send the command to controller
            //Return:
            //****************************************************************
            if (SCPort.IsOpen)
            {
                SCPort.Write(CommandString);
                SCPort.DiscardOutBuffer();
            }
        }

        public void Delay(long milliSecond = 500)
        {
            //****************************************************************
            //Function: Delay
            //Parameters: milliSecond:the waiting time, unit is millsecond
            //Description: appoint the waiting time and exit waiting until the data reading is finished or clicking the stop button or close the window or the waiting time is over.
            //Return:
            //****************************************************************
            int start = Environment.TickCount;

            BlnReadCom = false;
            BlnStopCommand = false;
            while (Math.Abs(Environment.TickCount - start) < milliSecond)
            {
                if (BlnReadCom == true)
                {
                    BlnReadCom = false;
                    return;
                }
                if (BlnStopCommand == true) return;
                Application.DoEvents();
            }
        }

        private void button1_Click(object sender, EventArgs e)      //Connect to appointed serial port
        {
            ConnectPort(Convert.ToInt16(textBox1.Text));
        }

        private void button2_Click(object sender, EventArgs e)      //Set new speed and get the current speed
        {
            sSpeed = Convert.ToInt32(textBox2.Text);
            if (sSpeed < 0 || sSpeed > 20)
            {
                MessageBox.Show("The speed value must be an integer between 0 and 20.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (BlnBusy == true)
            {
                MessageBox.Show("The connection is busy, please wait.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            sSpeed = Convert.ToInt32(sSpeed * 12.75);
            StrReceiver = "";
            BlnBusy = true;
            BlnSet = true;
            SendCommand("V" + sSpeed.ToString() + "\r");            //Set speed
            Delay(100000);
            BlnBusy = false;

            StrReceiver = "";
            BlnBusy = true;
            BlnSet = false;
            SendCommand("?V\r");                                    //Inquiry speed
            Delay(100000);
            BlnBusy = false;

            if (StrReceiver != "")
            {
                sSpeed = Convert.ToInt32(System.Text.RegularExpressions.Regex.Replace(StrReceiver, @"[^0-9]+", "")); //В указанной входной строке заменяет все строки, соответствующие указанному регулярному выражению, указанной строкой замены.
                label3.Text = "The speed is " + sSpeed.ToString();
            }
        }

        private void button4_Click(object sender, EventArgs e)      //Move all axis to the appointed position and get the current position
        {
            long lStep;
            string s;
            button7.Focus();
            button6_Click(sender, e);
            lStep = Convert.ToInt64(((Convert.ToDouble(textBox8.Text) * 200) - dCurrPosiX) / DblPulseEqui);
            if (lStep > 0)
                s = "+" + lStep.ToString();
            else
                s = lStep.ToString();
            StrReceiver = "";
            BlnBusy = true;
            BlnSet = true;
            SendCommand("X" + s + "\r");   //Move X axis to the appointed position.

            textBox11.Text = "...... ";
            timer1.Interval = 310 - Convert.ToInt32(sSpeed);
            timer1.Enabled = true;
            Delay(100000000);
            BlnBusy = false;
            timer1.Enabled = false;

            button7.Focus();
            button6_Click(sender, e);
            lStep = Convert.ToInt64(((Convert.ToDouble(textBox3.Text) * 200) - dCurrPosiY) / DblPulseEqui);
            if (lStep > 0)
                s = "+" + lStep.ToString();
            else
                s = lStep.ToString();
            StrReceiver = "";
            BlnBusy = true;
            BlnSet = true;
            SendCommand("Y" + s + "\r");   //Move X axis to the appointed position.

            textBox9.Text = "...... ";
            timer1.Interval = 310 - Convert.ToInt32(sSpeed);
            timer1.Enabled = true;
            Delay(100000000);
            BlnBusy = false;
            timer1.Enabled = false;
            button6_Click(sender, e);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            string s;
            if (BlnReadCom == true)
            {
                timer1.Enabled = false;
                return;
            }
            s = textBox11.Text;
            textBox11.Text = s.Substring(s.Length - 1, 1) + s.Substring(0, s.Length - 1);
        }

        private void button6_Click(object sender, EventArgs e)      //Get the current position of X axis
        {
            StrReceiver = "";
            BlnBusy = true;
            BlnSet = false;
            SendCommand("?X\r");            //Inquiry the current position of X axis
            Delay(100000);
            BlnBusy = false;

            if (StrReceiver != "")
            {
                if (StrReceiver.Substring(5, 1) == "-")
                    lCurrStep = -Convert.ToInt64(System.Text.RegularExpressions.Regex.Replace(StrReceiver, @"[^0-9]+", ""));
                else
                    lCurrStep = Convert.ToInt64(System.Text.RegularExpressions.Regex.Replace(StrReceiver, @"[^0-9]+", ""));
            }
            else
                return;
            dCurrPosiX = lCurrStep * DblPulseEqui;
            textBox11.Text = Convert.ToString(dCurrPosiX / 200);

            StrReceiver = "";
            BlnBusy = true;
            BlnSet = false;
            SendCommand("?Y\r");            //Inquiry the current position of Y axis
            Delay(100000);
            BlnBusy = false;

            if (StrReceiver != "")
            {
                if (StrReceiver.Substring(5, 1) == "-")
                    lCurrStep = -Convert.ToInt64(System.Text.RegularExpressions.Regex.Replace(StrReceiver, @"[^0-9]+", ""));
                else
                    lCurrStep = Convert.ToInt64(System.Text.RegularExpressions.Regex.Replace(StrReceiver, @"[^0-9]+", ""));
            }
            else
                return;
            dCurrPosiY = lCurrStep * DblPulseEqui;
            textBox9.Text = Convert.ToString(dCurrPosiY / 200);
        }

        private void button5_Click(object sender, EventArgs e)      //Return to origin
        {
            string HomeZ;
            double zStep;
            zStep = (Convert.ToDouble(textBox4.Text) * 200);
            if (zStep > 300)
                zStep = 300;
            HomeZ = "+" + zStep.ToString();

            button7.Focus();
            StrReceiver = "";
            BlnBusy = true;
            BlnSet = true;
            SendCommand("HX0\r");   //Home X axis

            textBox11.Text = "...... ";
            timer1.Interval = 310 - Convert.ToInt32(sSpeed);
            timer1.Enabled = true;
            Delay(1000000);
            timer1.Enabled = false;
            BlnBusy = false;
            button6_Click(sender, e);

            button7.Focus();
            StrReceiver = "";
            BlnBusy = true;
            BlnSet = true;
            SendCommand("HY0\r");   //Home Y axis

            textBox9.Text = "...... ";
            timer1.Interval = 310 - Convert.ToInt32(sSpeed);
            timer1.Enabled = true;
            Delay(1000000);
            timer1.Enabled = false;
            BlnBusy = false;
            button6_Click(sender, e);

            button7.Focus();
            StrReceiver = "";
            BlnBusy = true;
            BlnSet = true;
            SendCommand("HZ0\r");   //Home Z axis

            textBox9.Text = "...... ";
            timer1.Interval = 310 - Convert.ToInt32(sSpeed);
            timer1.Enabled = true;
            Delay(1000000);
            timer1.Enabled = false;
            BlnBusy = false;
            button6_Click(sender, e);

            button7.Focus();
            StrReceiver = "";
            BlnBusy = true;
            BlnSet = true;
            SendCommand("Z"+ HomeZ +"\r");   //Home Z axis

            textBox9.Text = "...... ";
            timer1.Interval = 310 - Convert.ToInt32(sSpeed);
            timer1.Enabled = true;
            Delay(1000000);
            timer1.Enabled = false;
            BlnBusy = false;
            button6_Click(sender, e);

        }

        private void button7_Click(object sender, EventArgs e)      //Stop moving
        {
            StrReceiver = "";
            BlnBusy = true;
            BlnSet = true;
            SendCommand("S\r");   //Stop moving
            Delay(100000000);
            timer1.Enabled = false;
            BlnStopCommand = true;
            //DelayWait(500);
            BlnBusy = false;

        }

        private void button8_Click(object sender, EventArgs e)
        {
            long lStep;
            double yStep;
            string x;
            string o;
            string y;
            long cycles;
            button7.Focus();
            button6_Click(sender, e);
            yStep = ((Convert.ToDouble(textBox10.Text) * 200) - dCurrPosiX) / DblPulseEqui;
            if (yStep > 0)
                x = "+" + yStep.ToString();
            else
                x = yStep.ToString();


            yStep = Convert.ToDouble(textBox7.Text) *200 / DblPulseEqui;
            if (yStep > 0)
                y = "+" + yStep.ToString();
            else
                y = yStep.ToString();

            cycles = Convert.ToInt64(Convert.ToDouble(textBox12.Text));
            o = Convert.ToString(-1 * Convert.ToInt64(x));

            for (; cycles > 0; cycles--)
            {

                StrReceiver = "";
                BlnBusy = true;
                BlnSet = true;
                SendCommand("X" + x + "\r");   //Move X axis to the appointed position.

                textBox11.Text = "......";
                timer1.Interval = 310 - Convert.ToInt32(sSpeed);
                timer1.Enabled = true;
                Delay(100000000);
                //BlnBusy = false;
                //timer1.Enabled = false;
                //button6_Click(sender, e);

                StrReceiver = "";
                //BlnBusy = true;
                //BlnSet = true;
                SendCommand("Y" + y + "\r");   //Move Y axis to the appointed position.

                textBox9.Text = "......";
                //timer1.Interval = 310 - Convert.ToInt32(sSpeed);
                //timer1.Enabled = true;
                Delay(100000000);
                //BlnBusy = false;
                ///timer1.Enabled = false;
                //button6_Click(sender, e);

                StrReceiver = "";
                //BlnBusy = true;
                //BlnSet = true;
                SendCommand("X" + o + "\r");   //Move Y axis to the appointed position.

                //textBox9.Text = "......";
                //timer1.Interval = 310 - Convert.ToInt32(sSpeed);
                //timer1.Enabled = true;
                Delay(100000000);
                //BlnBusy = false;
                //timer1.Enabled = false;
                //button6_Click(sender, e);

                StrReceiver = "";
                //BlnBusy = true;
                //BlnSet = true;
                SendCommand("Y" + y + "\r");   //Move Y axis to the appointed position.

                //textBox9.Text = "......";
                //timer1.Interval = 310 - Convert.ToInt32(sSpeed);
                //timer1.Enabled = true;
                Delay(100000000);
                BlnBusy = false;
                timer1.Enabled = false;
                //button6_Click(sender, e);
            }
            button6_Click(sender, e);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
