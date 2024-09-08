using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GemBox.Spreadsheet;
using System.IO.Ports;
using System.IO;


namespace cobak_ngeload_data
{
    public partial class Form1 : Form
    {
        string[] capturelog = new string[999999];
        int sheet_;
        int[] tertimbang = new int[999999];
        SerialPort port = new SerialPort();
        string RxString;
        //string ArduinoData = null;
        string[] ports = SerialPort.GetPortNames();
        int itungcom;
        bool adadata = false;
        bool cekdatatxt = false;
        string[] datatxt = new string[2];
        public Form1()
        {
            InitializeComponent();
            foreach (string s in ports) comboBox2.Items.Add(s);
        }


        ExcelFile ef, exportdata;        
        int penghitungsheet, idlog1,idlog2,idlog3;       Random rnd = new Random();
        bool sudahinput = false;
        bool notqa = false;



        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox4.Visible = false;
            button6.Enabled = false;
            comboBox3.Visible = false;
            label13.Visible = false;
            textBox4.Visible =false;            
            button1.Enabled = false;
            button5.Enabled = false;
            label5.Visible = false;
            label9.Visible = false;
            textBox2.Enabled = false;
            button4.Enabled = false;
            idlog1 = rnd.Next(1,1000);
            idlog2 = rnd.Next(1, 1000);
            idlog3 = rnd.Next(1, 1000);
            textBox3.PasswordChar = '*';
            SpreadsheetInfo.SetLicense("EQU2-1000-0000-000U");
            ef = ExcelFile.Load(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
            
            //ExcelWorksheet ws = ExcelWorksheet ws = ef.Worksheets[0];


            exportdata = ExcelFile.Load("templatedata.xlsx");
            penghitungsheet = ef.Worksheets.Count;

            #region PENGHITUNG SHEET DAN PENAMPIL KE COMBOBOX
            for (int a = 0; a < penghitungsheet; a++)
            {
                comboBox1.Items.Add(ef.Worksheets[a].Name);
                comboBox3.Items.Add(ef.Worksheets[a].Name); ///////UPDATED

            }
            #endregion
            textBox1.Enabled = false; //NONAKTIF TEXTBOX STANDART
            panel2.Visible = false;panel4.Visible = false;
            ef.Save(@"C:\DATA PENIMBANGAN\BACKUP\backup-of-tmplt-harian"+idlog1+idlog2+idlog3+ ".xlsx");

            timer2.Enabled = true;
            comboBox2.Visible = false;
            itungcom = comboBox2.Items.Count;
            //MessageBox.Show(itungcom.ToString());
            if (itungcom == 0)
            {
                timer2.Enabled = false;
            }

        }

        #region PASSWORD DAN EDIT DATA UNTUK DISAVE KE TEMPLATE EXCEL
        double  standar,min, max, toleransi,aktual;
        double[] dataitung = new double[4];
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            #region MENAMPILKAN STANDART KE TEXTBOX 1 DENGAN PARAMETER CARTON NUMBER
            try
            {
                if (textBox2.TextLength > 0)
                {
                    textBox1.Text = (Convert.ToString(ef.Worksheets[comboBox1.SelectedItem.ToString()].Cells[Convert.ToInt32(textBox2.Text) + 1, 2].Value));

                    if (textBox1.TextLength > 0&& textBox1.Text != "Standart (kg)")
                    {
                        String[] data = new String[4];

                        for (int aa = 0; aa < 4; aa++)
                        {
                            data[aa] = (Convert.ToString(ef.Worksheets[comboBox1.SelectedItem.ToString()].Cells[Convert.ToInt32(textBox2.Text) + 1, (aa + 2)].Value));
                            if (data[aa].Length <= 0)
                            {
                                data[aa] = "-";
                            }

                            if (data[aa] != "-")
                            {
                                dataitung[aa] = Convert.ToDouble(data[aa]);
                                if (aa == 3)
                                {
                                    dataitung[aa] = dataitung[aa] / 100 * dataitung[0];
                                }
                            }
                            else
                            {
                                dataitung[aa] = -1;
                            }
                        }
                        if(adadata == true)
                        {
                            button4.Enabled = true;
                        }
                        

                    }
                    else {
                        button4.Enabled = false;
                        textBox1.ResetText();

                    }

                    //////textBox1.Text = (Convert.ToString(ef.Worksheets[comboBox1.SelectedItem.ToString()].Cells[Convert.ToInt32(textBox2.Text) + 1, 2].Value));
                    //label5.Text = (Convert.ToString(ef.Worksheets[comboBox1.SelectedIndex].Cells[Convert.ToInt32(textBox2.Text) + 1, 2].Value));


                    ////ambil data
                    //if ((Convert.ToString(ef.Worksheets[comboBox1.SelectedIndex].Cells[Convert.ToInt32(textBox2.Text) + 1, 2].Value)) != "" || (Convert.ToString(ef.Worksheets[comboBox1.SelectedIndex].Cells[Convert.ToInt32(textBox2.Text) + 1, 3].Value)) != "" || (Convert.ToString(ef.Worksheets[comboBox1.SelectedIndex].Cells[Convert.ToInt32(textBox2.Text) + 1, 4].Value)) != "" || (Convert.ToString(ef.Worksheets[comboBox1.SelectedIndex].Cells[Convert.ToInt32(textBox2.Text) + 1, 5].Value)) != "")
                    //{
                    //    standar = Convert.ToDouble((Convert.ToString(ef.Worksheets[comboBox1.SelectedIndex].Cells[Convert.ToInt32(textBox2.Text) + 1, 2].Value)));
                    //    min = Convert.ToDouble((Convert.ToString(ef.Worksheets[comboBox1.SelectedIndex].Cells[Convert.ToInt32(textBox2.Text) + 1, 3].Value)));
                    //    max = Convert.ToDouble((Convert.ToString(ef.Worksheets[comboBox1.SelectedIndex].Cells[Convert.ToInt32(textBox2.Text) + 1, 4].Value)));
                    //    toleransi = Convert.ToDouble((Convert.ToString(ef.Worksheets[comboBox1.SelectedIndex].Cells[Convert.ToInt32(textBox2.Text) + 1, 5].Value)));
                    //    toleransi = (toleransi/100)*standar;

                    label9.Text = ("standar : " + dataitung[0].ToString() + " min : " + dataitung[1].ToString() + " max : " + dataitung[2].ToString() + " tol : " + dataitung[3].ToString());
                    //}
                    //else
                    //{
                    //    MessageBox.Show("data carton No tidak ada!");


                    //}
                    ////end of ambil data

                }
                else
                {
                    button4.Enabled = false;
                    textBox1.ResetText();
                }
            }
            catch { }
            #endregion


            ////cek data sudah input atau belum
            //try
            //{
            //    exportdata = ExcelFile.Load(@"C:\DATA PENIMBANGAN\LOG CAPTURE\" + comboBox1.SelectedItem.ToString() + " - " + idlog1 + idlog2 + idlog3 + ".xlsx");
            //    String cektgl = (exportdata.Worksheets[0].Cells[Convert.ToInt32(textBox2.Text) + 1, 8].Value).ToString();
            //    if (cektgl.Length > 0)
            //    {
            //        sudahinput = true;
            //    }
            //    else
            //    {
            //        sudahinput = false;
            //    }
            //}

            //catch
            //{

            //}
            

            //if (sudahinput == true)
            //{
               
            //    MessageBox.Show("Carton No. " +textBox2.Text+ " already inserted", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    sudahinput = false;
            //    textBox2.ResetText();
            //}
            ////end of check data

        }

        
        void itung() {

            standar = dataitung[0];
            min = dataitung[1];
            max = dataitung[2];
            toleransi  = dataitung[3];
            
            if (textBox4.TextLength > 0&&textBox2.TextLength>0)
            {
                try { aktual = Convert.ToDouble(textBox4.Text); } catch { }
                


                if (standar == -1 && min == -1 && max == -1 && toleransi == -1)
                {
                    label8.Text = "UNKNOWN";
                    notqa = true;
                    return;
                }
                else if(standar == -1 && min == -1 && max == -1 )
                {
                    label8.Text = "UNKNOWN";
                    notqa = true;
                    return;
                }

                if (max >= min)
                {
                    if (aktual >= min && aktual <= max)
                    {
                        label8.Text = "QUALIFIED";
                        
                        notqa = false;
                        return;
                    }
                }
                else
                {
                    if (aktual >= min)
                    {
                        label8.Text = "QUALIFIED";
                        
                        notqa = false;
                        return;
                    }
                    else if (aktual <= max)
                    {
                        label8.Text = "QUALIFIED";
                        
                        notqa = false;
                        return;
                    }
                }

                if (aktual == standar)
                {
                    label8.Text = "QUALIFIED";
                    
                    notqa = false;
                    return;
                }
                else if (aktual <= (toleransi + standar) && aktual >= (standar - toleransi))
                {
                    label8.Text = "QUALIFIED";
                    
                    notqa = false;
                    return;
                }
                else
                {
                    label8.Text = "NOT QUALIFIED";
                    notqa = true;
                    return; 
                }
            }
        }


        bool loopgantieditsave = false;
        private void button1_Click(object sender, EventArgs e){
            button5.Enabled = false;
            try
            {
                if (cekdatatxt == false)
                {
                    exportdata = ExcelFile.Load(@"C:\DATA PENIMBANGAN\LOG CAPTURE\" + comboBox1.SelectedItem.ToString() + " - " + idlog1 + idlog2 + idlog3 + ".xlsx");
                }
                else {
                    exportdata = ExcelFile.Load(@"C:\DATA PENIMBANGAN\LOG CAPTURE\" + comboBox1.SelectedItem.ToString() + datatxt[0] + ".xlsx");
                }
               // exportdata = ExcelFile.Load(@"C:\DATA PENIMBANGAN\LOG CAPTURE\" + comboBox1.SelectedItem.ToString() + " - " + idlog1 + idlog2 + idlog3 + ".xlsx");
                String cektgl = (exportdata.Worksheets[0].Cells[Convert.ToInt32(textBox2.Text) + 1, 8].Value).ToString();
                if (cektgl.Length > 0)
                {
                    sudahinput = true;
                }
                else
                {
                    sudahinput = false;
                }
            }

            catch
            {

            }


            if (sudahinput == true)
            {
                MessageBox.Show("Cannot edit, Carton No. " + textBox2.Text + " already inserted", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                sudahinput = false;
                textBox2.ResetText();

            }
            else {
                if (loopgantieditsave == true)
                {
                    try
                    {
                        ef.Save(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("Silahkan tutup excel dahulu" + ex.ToString());
                        return;
                    }
                    button1.Text = "EDIT";
                    textBox1.Enabled = false; //AKTIFKAN TEXTBOX STANDART

                    loopgantieditsave = false;
                    ef.Worksheets[comboBox1.SelectedItem.ToString()].Rows[(Convert.ToInt32(textBox2.Text) + 1)].Cells[2].SetValue(textBox1.Text);
                    try
                    {

                        ef.Save(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
                        MessageBox.Show("Edit succes", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        textBox2.Enabled = true;
                        button4.Enabled = true;
                        button6.Enabled = true;
                        if (Convert.ToInt16(label7.Text) > 0)
                        {
                            button5.Enabled = true;
                        }
                        else
                        {

                            button5.Enabled = false;
                        }
                       
                        
                      
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("Silahkan tutup excel terlebih dahulu!" + ex.ToString());
                        //textBox1.Text = label5.Text;

                        return;
                    }


                }

                else
                {
                    textBox3.Visible = true;
                    button2.Visible = true;
                    button3.Visible = true;
                    label4.Visible = true;
                    button6.Enabled = false;
                    panel2.Visible = true;
                    panel4.Visible = true;
                    textBox2.Enabled = false;
                    button4.Enabled = false;
                    //button5.Enabled = false;
                    button1.Enabled = false;
                    //button5.Enabled = false;
                    textBox3.Clear();
                }
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {


            if (export == true)
            {

                if (textBox3.Text == "divisitimbang")
                {
                    export = false;
                    panel4.Visible = false;                    
                    panel2.Visible = false;
                    label4.Visible = false;
                    textBox3.Visible = false;
                    button2.Visible = false;
                    button3.Visible = false;

                    //gawe simpen exports

                    try
                    {
                        if (sheet_ == 1)
                        {
                            
                            ef.Worksheets.Add("Penimbangan Selesai");
                            ef.Worksheets.Remove(comboBox1.SelectedItem.ToString());
                            ef.Save(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
                            exportdata.Save(@"C:\DATA PENIMBANGAN\EXPORT\" + comboBox1.SelectedItem.ToString() + ".xlsx");
                            MessageBox.Show("Export " + comboBox1.SelectedItem + " succes", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            ef = ExcelFile.Load(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
                            button6.Enabled = false;
                            File.Delete(comboBox1.SelectedItem + ".txt");
                            File.Delete(comboBox1.SelectedItem + "log .txt");
                            button5.Enabled = false; //buat lock exports

                            comboBox1.ResetText();
                            comboBox1.Items.Clear();
                            penghitungsheet = ef.Worksheets.Count;
                            for (int a = 0; a < penghitungsheet; a++)
                            {
                                comboBox1.Items.Add(ef.Worksheets[a].Name);

                            }

                            GC.Collect();
                        }
                        else
                        {
                            ef.Worksheets.Remove(comboBox1.SelectedItem.ToString());
                            ef.Save(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
                            exportdata.Save(@"C:\DATA PENIMBANGAN\EXPORT\" + comboBox1.SelectedItem.ToString() + ".xlsx");
                            MessageBox.Show("Export " + comboBox1.SelectedItem + " succes", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            ef = ExcelFile.Load(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
                            button6.Enabled = false;
                            File.Delete(comboBox1.SelectedItem + ".txt");
                            File.Delete(comboBox1.SelectedItem + "log .txt");
                            button5.Enabled = false; //buat lock exports

                            comboBox1.ResetText();
                            comboBox1.Items.Clear();
                            penghitungsheet = ef.Worksheets.Count;
                            for (int a = 0; a < penghitungsheet; a++)
                            {
                                comboBox1.Items.Add(ef.Worksheets[a].Name);

                            }
                            GC.Collect();
                        }


                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("Tutup Excel terlebih dahulu"+ex.ToString());
                        return;
                    }

                   // comboBox1.Items.Clear();
                    //comboBox1.ResetText();
                    label6.Text = "0";
                    label7.Text = "0";
                    button1.Enabled = false; //button5.Enabled = false;
                    comboBox1.Enabled = true;
                    textBox3.ResetText();


                    comboBox1.Items.Remove(comboBox1.SelectedItem);
                    //penghitungsheet = ef.Worksheets.Count;
                    //for (int a = 0; a < penghitungsheet; a++)
                    //{
                    //    comboBox1.Items.Add(ef.Worksheets[a].Name);

                    //}
                    //gawe simpen exports




                }
                else
                {
                    MessageBox.Show("Wrong Password", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBox3.ResetText();
                }
            }
            else {
                if (textBox3.Text == "divisitimbang")
                {
                    panel2.Visible = false; panel4.Visible = false;
                    button1.Enabled = true;
                    // button5.Enabled = false;
                    button1.Text = "SAVE";
                    textBox1.Enabled = true; //AKTIFKAN TEXTBOX STANDART
                    loopgantieditsave = true;
                    button6.Enabled = false;
                    textBox3.ResetText();


                }
                else
                {
                    MessageBox.Show("Wrong Password", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBox3.ResetText();
                }
            }

            
        }

        


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            comboBox4.Items.Clear();
            
            button6.Enabled = true;
            comboBox3.SelectedItem = comboBox1.SelectedItem; //UPDATED
            if (comboBox1.Text == "Penimbangan Selesai") {
                DialogResult result = MessageBox.Show("Penimbangan Selesai, Exit ?", "");


                if (result == DialogResult.Yes) {
                    Environment.Exit(0);
                    
                }
            }

            button4.Enabled = false;
            //button5.Enabled = true;
            textBox2.Enabled = true;
            try {
                ef = ExcelFile.Load(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
                exportdata = ExcelFile.Load("templatedata.xlsx");
                var usedRange = ef.Worksheets[comboBox1.SelectedItem.ToString()].GetUsedCellRange(true);
                int lastUsedRow = usedRange.LastRowIndex;
                label6.Text = (lastUsedRow - 1).ToString();
                comboBox1.Enabled = false;
                //data = 0;
                label7.Text = data.ToString();
                
            }
            catch {
                Environment.Exit(0);
            }

            //if (tertimbang[comboBox3.SelectedIndex] == 0)
            //{
            //    button5.Enabled = false;
            //}
            //else {
            //    button5.Enabled = true;
            //}
            if (data == Convert.ToInt16(label6.Text))
            {
                button5.Enabled = true;
            }
            else
            {
               // button5.Enabled = false;
            }



            //gawe cek jumlah data tertimbang            
            try
            {

                String fileName = comboBox1.SelectedItem + ".txt";
                StreamReader perintah = File.OpenText(fileName);
                String baca;
                if ((baca = perintah.ReadLine()) != null)
                {
                    datatxt = baca.Split('#');
                    label7.Text = datatxt[1];
                    cekdatatxt = true;
                }
                else
                {
                    label7.Text = "0";
                }
                perintah.Close();
            }
            catch
            {
                cekdatatxt = false;
            }
            //gawe cek jumlah data tertimbang           
            




            try
            {
                String[] loggertimbang = new string[999999];
                String fileName = comboBox1.SelectedItem + "log .txt";
                StreamReader perintah2 = File.OpenText(fileName);
                String moco;
                if ((moco = perintah2.ReadLine()) != null)
                {                    

                    loggertimbang = moco.Split(' ');
                    for (int se = 1; se <= (loggertimbang.Length - 1); se++) {
                        comboBox4.Items.Add(loggertimbang[se-1]);
                    }
                   
                    perintah2.Close();

                }
                else
                {
                    
                }
               
            }
            catch
            {
                //logging yg tertimbang
                for (int tt = 1; tt <= Convert.ToInt16(label6.Text); tt++)
                {
                    comboBox4.Items.Add(tt);
                    String fileName = comboBox1.SelectedItem + "log .txt";
                    //gawe overwrite data tertimbang            
                    StreamWriter sw = new StreamWriter(fileName,true);
                    sw.Write(tt + " ");
                    sw.Close();
                }                
            }


            //gawe cek data yg tertimbang


            //gawe cek data yg tertimbang



            if (Convert.ToInt16(label7.Text) > 0)
            {
                button5.Enabled = true;
            }
            else {
                button5.Enabled = false;
            }
            if (textBox2.TextLength > 0)
            {
                button1.Enabled = true;
                button4.Enabled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            kuncitombolonlyangka(sender,e);
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            //if (textBox4.TextLength > 0 && textBox1.TextLength > 0)
            //{
            //    itung();
            //}
            //else {
            //    label8.Text = "-";
            //}
        }

        private void comboBox2_Click(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            ports = SerialPort.GetPortNames();
            foreach (string s in ports) comboBox2.Items.Add(s);
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            
            itungcom--;

            // comPort = Convert.ToString(comboBox2.SelectedItem); // Set selected COM port
            try
            {
                port.Close();
                port.PortName = Convert.ToString(comboBox2.Items[itungcom]);
                port.BaudRate = 9600;
                //port.DataBits = 8;
                //port.Parity = Parity.None;
               // port.StopBits = StopBits.One;
                port.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(serialPort1_DataReceived);
                port.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to Connect :: " + ex.Message, "Please Check the Hardware", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }

            if (itungcom == 0)
            {
                timer2.Enabled = false;
                
            }
            
        }

        int stopp = 8;
        private void label1_TextChanged(object sender, EventArgs e)
        {
            if (stopp == 8)
            {
                timer2.Enabled = false;
            }
            stopp = 1;
            textBox4.Text = label1.Text;
        }

        delegate void dataku(String d);

        void datamasuk(String data)
        {
            label1.Text = data.ToString();
            if (textBox4.TextLength > 0 && textBox1.TextLength > 0)
            {
                itung();
            }
            else
            {
                label8.Text = "-";
            }


        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.TextLength > 0)
            {
                button1.Enabled = true;
            }
            else {
                button1.Enabled = false;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label11.Text = (Convert.ToString(DateTime.Now.ToShortDateString() + " - " + Convert.ToString(DateTime.Now.ToLongTimeString())));
            GC.Collect();

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            var result = MessageBox.Show("Exit", "",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

            e.Cancel = (result == DialogResult.No);
            overwritelog();
            
        }

        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            if (port.IsOpen == true)
            {
                try
                {

                    RxString = port.ReadLine();
                    //richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(RxString)));
                    //richTextBox1.ScrollToCaret();
                    this.Invoke(new dataku(datamasuk), RxString);
                    //Invalidate();
                    //richTextBox1.Refresh();
                    adadata = true;

                }
                catch { }
            }

        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button2_Click(sender, e);
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (port.IsOpen == true)
            {

            }
            else {
                MessageBox.Show("data ilang");
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
           // backgroundWorker1.RunWorkerAsync();
        }

         void oke()
        {
            System.Media.SoundPlayer player = new System.Media.SoundPlayer("alert.wav");
            player.Play();
        }
        void gagal()
        {
            System.Media.SoundPlayer player = new System.Media.SoundPlayer("beep.wav");
            player.Play();
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            gagal();
        }

        private void label8_TextChanged(object sender, EventArgs e)
        {
            if (label8.Text == "QUALIFIED")
            {
                label13.Text = "oke";
                oke();

                
            }
            else {
                label13.Text = "gakoke";
            }
        }

        void overwritelog() {
            if (Convert.ToInt16(label7.Text) < Convert.ToInt16(label6.Text)) {
                try
                {
                    String fileName = comboBox1.SelectedItem + ".txt";
                    //gawe overwrite data tertimbang            
                    StreamWriter sw = new StreamWriter(fileName);
                    if (cekdatatxt == false)
                    {
                        
                        sw.Write(" - " + idlog1 + idlog2 + idlog3 + "#" + label7.Text);

                    }
                    else {
                        
                        sw.Write(datatxt[0] + "#" + label7.Text);
                    }
                        sw.Close();


                    //gwer overwrite data tertimbang
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(Ex.ToString());
                }
            }

           
        }

        private void button6_Click(object sender, EventArgs e)
        {
           
            overwritelog();
            cekdatatxt = false;
           // button5.Enabled = false;
            label6.Text = "0";
            label7.Text = "0";
            textBox2.Enabled = false; //disable carton no
            comboBox1.Enabled = true;
            button6.Enabled = false;
            button1.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
        }


        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            data = tertimbang[comboBox3.SelectedIndex];
            label7.Text = data.ToString();

        }

        private void button8_Click(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            comboBox1.FindString(comboBox1.Text);
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            button4.Enabled = true;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            overwritelog();
        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            kuncitombol(sender, e);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
          
            
            

            //if ((baca = perintah.ReadLine()) != null)
            //{
               
            //    capturelog = baca.Split('#');
            //    int panjangarray = capturelog.Length-1;
            //    bool adadata_ = false;
            //    MessageBox.Show(capturelog[0]);
               
            //    for (int cc=1; cc <= panjangarray; cc++)
            //    {
            //        for (int xx = 1; xx <= Convert.ToInt16(label6.Text); xx++) {
            //            if (Convert.ToInt16(capturelog[cc-1]) == xx)
            //            {
            //                adadata_ = false;
            //            }
            //            else
            //            {
            //                adadata_ = true;
            //                break;
            //            }
            //        }

            //        if (adadata_ == true)
            //        {
            //            comboBox4.Items.Add(cc);
            //            adadata_ = false;
            //        }

            //    }
            //}
            //else
            //{
                
            //}
            //perintah.Close();

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click_1(object sender, EventArgs e)
        {

            //int[] numbers = { 1, 3, 4, 9, 2 };
            //int numToRemove = 4;
            //numbers = numbers.Where(val => val != numToRemove).ToArray();


            //comboBox4.Items.RemoveAt(comboBox4.Items.IndexOf(textBox2.Text));

            //MessageBox.Show(comboBox4.Items.IndexOf(.ToString());

            //comboBox4.Items.Clear();
            //comboBox4.Refresh();
            
        }

        private void label1_Click(object sender, EventArgs e)
        {
            oke();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Enabled = true;
            button4.Enabled = true;
            panel2.Visible = false;panel4.Visible = false;
            button1.Enabled = true;           
            button6.Enabled = true;
            if (Convert.ToInt16(label7.Text) > 0)
            {
                button5.Enabled = true;
            }
            else
            {

                button5.Enabled = false;
            }
            return;
        }
        #endregion

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (textBox2.TextLength > 0 && textBox1.Text!= "Standart (kg)" && textBox1.TextLength>0&& adadata==true && label8.Text!="-") {
                if (e.KeyCode == Keys.Enter)
                {
                    button4_Click(sender,e);
                }
            }

        }

        int data = 0;
        
        private void button4_Click(object sender, EventArgs e)
        {
            
            if (notqa==true)
            {
              //  try { port.Close(); } catch { }
               gagal();
                timer3.Enabled = true;
                
                
                // MessageBox.Show("Recheck your carton");
                
                DialogResult result = MessageBox.Show("Recheck your carton", "", MessageBoxButtons.OK);
                //port.Close();
                if (result == DialogResult.OK)
                {
                   // port.Open();
                    timer3.Enabled = false;
                }
                
            }
            else
            {

                //cek data sudah input atau belum
                try
                {
                    if (cekdatatxt == false)
                    {
                        exportdata = ExcelFile.Load(@"C:\DATA PENIMBANGAN\LOG CAPTURE\" + comboBox1.SelectedItem.ToString() + " - " + idlog1 + idlog2 + idlog3 + ".xlsx");
                    }
                    else {
                        exportdata = ExcelFile.Load(@"C:\DATA PENIMBANGAN\LOG CAPTURE\" + comboBox1.SelectedItem.ToString() + datatxt[0] + ".xlsx");
                    }

                   
                    String cektgl = (exportdata.Worksheets[0].Cells[Convert.ToInt32(textBox2.Text) + 1, 8].Value).ToString();
                    if (cektgl.Length > 0)
                    {
                        sudahinput = true;
                    }
                    else
                    {
                        sudahinput = false;
                    }
                }

                catch
                {

                }


                if (sudahinput == true)
                {

                    MessageBox.Show("Carton No. " + textBox2.Text + " already inserted", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    sudahinput = false;
                    textBox2.ResetText();
                }
                //end of check data
                else
                {
                    try
                    {
                        exportdata.Worksheets[0].Rows[(Convert.ToInt32(textBox2.Text) + 1)].Cells[0].SetValue(comboBox1.SelectedItem.ToString());
                        exportdata.Worksheets[0].Rows[(Convert.ToInt32(textBox2.Text) + 1)].Cells[1].SetValue(textBox2.Text);
                        exportdata.Worksheets[0].Rows[(Convert.ToInt32(textBox2.Text) + 1)].Cells[2].SetValue((Convert.ToString(ef.Worksheets[comboBox1.SelectedItem.ToString()].Cells[Convert.ToInt32(textBox2.Text) + 1, 2].Value)));
                        exportdata.Worksheets[0].Rows[(Convert.ToInt32(textBox2.Text) + 1)].Cells[3].SetValue((Convert.ToString(ef.Worksheets[comboBox1.SelectedItem.ToString()].Cells[Convert.ToInt32(textBox2.Text) + 1, 3].Value)));
                        exportdata.Worksheets[0].Rows[(Convert.ToInt32(textBox2.Text) + 1)].Cells[4].SetValue((Convert.ToString(ef.Worksheets[comboBox1.SelectedItem.ToString()].Cells[Convert.ToInt32(textBox2.Text) + 1, 4].Value)));
                        exportdata.Worksheets[0].Rows[(Convert.ToInt32(textBox2.Text) + 1)].Cells[5].SetValue((Convert.ToString(ef.Worksheets[comboBox1.SelectedItem.ToString()].Cells[Convert.ToInt32(textBox2.Text) + 1, 5].Value)));
                        exportdata.Worksheets[0].Rows[(Convert.ToInt32(textBox2.Text) + 1)].Cells[6].SetValue(label1.Text);
                        exportdata.Worksheets[0].Rows[(Convert.ToInt32(textBox2.Text) + 1)].Cells[7].SetValue(label8.Text);
                        exportdata.Worksheets[0].Rows[(Convert.ToInt32(textBox2.Text) + 1)].Cells[8].SetValue(Convert.ToString(DateTime.Now.ToShortDateString() + " - " + Convert.ToString(DateTime.Now.ToLongTimeString())));
                        // exportdata.Worksheets[0].Name = (Convert.ToString(comboBox1.SelectedItem));

                        if (cekdatatxt == false)
                        {
                            exportdata.Save(@"C:\DATA PENIMBANGAN\LOG CAPTURE\" + comboBox1.SelectedItem.ToString() + " - " + idlog1 + idlog2 + idlog3 + ".xlsx");

                        }
                        else {
                            exportdata.Save(@"C:\DATA PENIMBANGAN\LOG CAPTURE\" + comboBox1.SelectedItem.ToString() + datatxt[0] + ".xlsx");
                        }

                         MessageBox.Show("Capture Succes : " + label1.Text + " kg" + " at " + Convert.ToString(DateTime.Now.ToShortDateString() + " - " + Convert.ToString(DateTime.Now.ToLongTimeString())));




                        comboBox4.SelectedIndex = comboBox4.FindString(textBox2.Text);
                        comboBox4.Items.Remove(comboBox4.SelectedItem);

                        //logging yg tertimbang


                        String fileName = comboBox1.SelectedItem + "log .txt";
                        //gawe overwrite data tertimbang            
                        StreamWriter sw = new StreamWriter(fileName);
                        int jmlcm = comboBox4.Items.Count;
                        for (int uye = 1; uye <= jmlcm; uye++) {

                            comboBox4.SelectedIndex = (uye - 1); 
                            sw.Write(comboBox4.SelectedItem+" ");
                        }
                                                
                        sw.Close();
                        //logging yg tertimbang



                        tertimbang[comboBox3.SelectedIndex]++;
                        data = tertimbang[comboBox3.SelectedIndex];
                        //data tersimpan  label7.Text = data.ToString();
                        label7.Text = (Convert.ToInt16(label7.Text) + 1).ToString();
                        sudahinput = false;
                        textBox2.ResetText();

                        if (data == Convert.ToInt16(label6.Text))
                        {
                            button5.Enabled = true;
                        }
                        else {
                          //  button5.Enabled = false;
                        }
                        
                        
                       // label14.Text = tertimbang[comboBox3.SelectedIndex].ToString();

                    }
                    catch (Exception ex){ MessageBox.Show("Silahkan tutup excel dahulu !"+ex.ToString()); return; }

                }



                try {
                    if (cekdatatxt == false)
                    {
                        exportdata.Save(@"C:\DATA PENIMBANGAN\LOG CAPTURE\" + comboBox1.SelectedItem.ToString() + " - " + idlog1 + idlog2 + idlog3 + ".xlsx");
                    }
                    else {
                        exportdata.Save(@"C:\DATA PENIMBANGAN\LOG CAPTURE\" + comboBox1.SelectedItem.ToString() + datatxt[0] + ".xlsx");
                    }


                   
                }

                catch(Exception ex){
                    MessageBox.Show("Silahkan tutup excel dahulu !"+ex.ToString()); return;
                }
                ////cek data sudah input atau belum
                //try
                //{
                //    exportdata = ExcelFile.Load(@"C:\Users\antares\Desktop\Log\" + comboBox1.SelectedItem.ToString() + " - " + idlog1 + idlog2 + idlog3 + ".xlsx");
                //    String cektgl = (exportdata.Worksheets[0].Cells[Convert.ToInt32(textBox2.Text) + 1, 8].Value).ToString();
                //    if (cektgl.Length > 0)
                //    {
                //        sudahinput = true;
                //    }
                //    else {
                //        sudahinput = false;
                //    }
                //}

                //catch{

                //}
                ////end of check data


                //if (sudahinput == true)
                //{
                //    MessageBox.Show("data sudah terinput");
                //    sudahinput = false;
                //}
                // else {

                // }

            }

            if (Convert.ToInt16(label7.Text) > 0)
            {
                button5.Enabled = true;
            }
            else
            {

                button5.Enabled = false;
            }
            overwritelog();
            //Application.Restart();
            //Environment.Exit(0);
            GC.Collect();
        }

        bool export = false;
        private void button5_Click(object sender, EventArgs e)
        {
            
            export = true;
            //cek jumlah sheet
            ef = ExcelFile.Load(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
            try {
                if (cekdatatxt == false)
                {
                    exportdata = ExcelFile.Load(@"C:\DATA PENIMBANGAN\LOG CAPTURE\" + comboBox1.SelectedItem.ToString() + " - " + idlog1 + idlog2 + idlog3 + ".xlsx");
                }
                else {
                    exportdata = ExcelFile.Load(@"C:\DATA PENIMBANGAN\LOG CAPTURE\" + comboBox1.SelectedItem.ToString() + datatxt[0] + ".xlsx");
                }
                }
            catch { MessageBox.Show("Gagal Export Data!");
                return;
            }
           
            //exportdata = ExcelFile.Load("templatedata.xlsx");
            
            sheet_ = ef.Worksheets.Count;
            //MessageBox.Show(sheet_.ToString());
            //check jumlah sheet

            try
            {
                ef.Save(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
            }
            catch(Exception ex) {
                MessageBox.Show("tutup excel dahulu"+ex.ToString());
                return;
            }


            var usedRange = ef.Worksheets[comboBox1.SelectedItem.ToString()].GetUsedCellRange(true);
            int lastUsedRow = usedRange.LastRowIndex;

            //if (data == (lastUsedRow - 1))
            if (Convert.ToInt16(label7.Text) == (lastUsedRow - 1))
            {
                try
                {
                    if (sheet_ == 1)
                    {
                        ef.Worksheets.Add("Penimbangan Selesai");
                        ef.Worksheets.Remove(comboBox1.SelectedItem.ToString());
                        ef.Save(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
                        exportdata.Save(@"C:\DATA PENIMBANGAN\EXPORT\" + comboBox1.SelectedItem.ToString() + ".xlsx");
                        MessageBox.Show("Export " + comboBox1.SelectedItem + " succes", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ef = ExcelFile.Load(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
                        button6.Enabled = false;
                        File.Delete(comboBox1.SelectedItem + ".txt");
                        File.Delete(comboBox1.SelectedItem + "log .txt");


                        button5.Enabled = false; //buat lock exports
                        comboBox1.Items.Remove(comboBox1.SelectedItem);
                        comboBox1.ResetText();
                        comboBox1.Items.Clear();
                        penghitungsheet = ef.Worksheets.Count;
                        for (int a = 0; a < penghitungsheet; a++)
                        {
                            comboBox1.Items.Add(ef.Worksheets[a].Name);

                        }
                        GC.Collect();

                    }
                    else
                    {
                        ef.Worksheets.Remove(comboBox1.SelectedItem.ToString());
                        ef.Save(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
                        exportdata.Save(@"C:\DATA PENIMBANGAN\EXPORT\" + comboBox1.SelectedItem.ToString() + ".xlsx");
                        MessageBox.Show("Export " + comboBox1.SelectedItem + " succes", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ef = ExcelFile.Load(@"C:\DATA PENIMBANGAN\TEMPLATE HARIAN\template harian.xlsx");
                        button6.Enabled = false;
                        File.Delete(comboBox1.SelectedItem + ".txt");
                        File.Delete(comboBox1.SelectedItem + "log .txt");

                        button5.Enabled = false; //buat lock exports
                        comboBox1.Items.Remove(comboBox1.SelectedItem);
                        comboBox1.ResetText();
                        comboBox1.Items.Clear();
                        penghitungsheet = ef.Worksheets.Count;
                        for (int a = 0; a < penghitungsheet; a++)
                        {
                            comboBox1.Items.Add(ef.Worksheets[a].Name);

                        }
                        GC.Collect();
                    }
                } catch( Exception ex){

                    MessageBox.Show("Tutup Excel terlebih dahulu" + ex.ToString());
                    return;
                }
                //comboBox1.Items.Clear();
               // comboBox1.ResetText();
                label6.Text = "0";
                label7.Text = "0";
                button1.Enabled = false; //button5.Enabled = false;
                comboBox1.Enabled = true;

            }
            else {
                //DialogResult dialogResult = MessageBox.Show("Jumlah data tertimbang tidak sesuai, " + comboBox1.SelectedItem + " akan terhapus. tetap export ?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                String fileNames = comboBox1.SelectedItem + "log .txt";
                StreamReader perintah3 = File.OpenText(fileNames);
                String mocos;
                mocos = perintah3.ReadLine();
                perintah3.Close();
                DialogResult dialogResult = MessageBox.Show("Karton : "+mocos+"belum tertimbang." +"\n"+"\n" + "Silahkan masukkan admin password!", "Konfirmasi Export", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Yes)
                {


                    //buat cek password
                    panel4.Visible = true;
                    panel2.Visible = true;
                    label4.Visible = true;
                    textBox3.Visible = true;
                    button2.Visible = true;
                    button3.Visible = true;                    
                    //buat cek password
                  





                }
                else
                {
                    return;
                }
            }


           
            //penghitungsheet = ef.Worksheets.Count;
            //MessageBox.Show(penghitungsheet.ToString());
            //comboBox1.Items.Clear();
            ////comboBox1.ResetText();
            //for (int a = 0; a < penghitungsheet; a++)
            //{
            //    comboBox1.Items.Add(ef.Worksheets[a].Name);

            //}
            //textBox2.Enabled = false;
            //button4.Enabled = false;
            //textBox2.ResetText();
            //textBox1.ResetText();
            //button5.Enabled = false;
        }


        #region kunci tombol huruf dan simbol
        void kuncitombol(object teksinput, KeyPressEventArgs angka)
        {

            if (!char.IsControl(angka.KeyChar) && !char.IsDigit(angka.KeyChar) && (angka.KeyChar != '.'))
            {
                angka.Handled = true;
            }

            // only allow one decimal point
            if ((angka.KeyChar == '.') && ((teksinput as TextBox).Text.IndexOf('.') > -1))
            {
                angka.Handled = true;
            }
        }

        void kuncitombolonlyangka(object teksinput2, KeyPressEventArgs angka2) {
            if (!char.IsControl(angka2.KeyChar) && !char.IsDigit(angka2.KeyChar))
            {
                angka2.Handled = true;
            }
        }
        #endregion
    }
}
