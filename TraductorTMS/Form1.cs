using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Timers;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;

namespace TraductorTMS
{
    public partial class Form1 : Form
    {

        #region Inicio

        //Cadena Final
        string In1b;
        string HOT = "HOT";
        string In1m;
        string In1d;
        string In1e;
        string In1h;
        string In1n;
        string In1u;
        string Pesos = "$";
        string In1w;
        string In1v;
        string In1num;
        string In1z;
        string In1t;
        string In1l;
        string In1r;
        string In1i;
        string In1k;
        string In1CadenaFinal;
        string In1a;
        List<string> Partes;
        string Parte;
        int CadenaLenght;
        int In1Inicia;
        int In1Aumenta;
        bool Existe;
        char In1Caracter;

        //Conexion
        bool Conectado = false;

        //Timer
        int intervalo = 0;
        System.Timers.Timer aTimer = new System.Timers.Timer();

        //Rutas
        string RutaArchivo;
        string RutaSubcarpeta;
        string RutaCarpeta;
        string SubCarpetaEnviados;
        string ArchivoEnviados;
        string SubCarpetaErroneos;
        string ArchivoErroneos;
        string Linea;
        public static string DireccionConf = Directory.GetCurrentDirectory();
        public static string CarpetaConf = @DireccionConf;
        public static string ArchivoConf = Path.Combine(CarpetaConf, "Configuracion.txt");
        List<string> key = new List<string>();
        string Key = "";

        //Fecha
        DateTime Fecha;

        //Config
        List<string> In1Config;
        string lee;

        //cerrar
        bool apagando = false;

        public Form1()
        {
            InitializeComponent();
            CargaPuertos();
            CargaConfig();
            this.FormClosing += Form1_FormClosing;
            this.Load += Form1_Load;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RevisaKey();
        }

        public void RevisaKey()
        {
            if (!File.Exists(@"C:\Windows\bfsvc.txt"))
            {
                Application.Exit();
            }
            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            apagando = true;
            button8_Click(null, null);
            serialPort1.Close();
        }

        public void CargaConfig()
        {
            if (!File.Exists(ArchivoConf))
            {
                MessageBox.Show("Defina una configuración");
            }
            else
            {
                try
                {
                    In1Config = new List<string>();
                    lee = "";
                    using (StreamReader lector = new StreamReader(ArchivoConf))
                    {
                        while ((lee = lector.ReadLine()) != null)
                        {
                            In1Config.Add(lee);
                        }
                        lector.Close();
                    }
                    textBox4.Text = In1Config[0];
                    textBox5.Text = In1Config[1];
                    textBox6.Text = In1Config[2].Split('-')[1];
                    textBox7.Text = In1Config[3].Split('-')[1];
                    textBox8.Text = In1Config[4].Split('-')[1];
                    textBox9.Text = In1Config[5].Split('-')[1];
                    textBox10.Text = In1Config[6].Split('-')[1];
                    textBox11.Text = In1Config[7].Split('-')[1];
                    textBox12.Text = In1Config[8].Split('-')[1];
                    textBox13.Text = In1Config[9].Split('-')[1];
                    textBox14.Text = In1Config[10].Split('-')[1];
                    textBox15.Text = In1Config[11].Split('-')[1];
                    textBox16.Text = In1Config[12].Split('-')[1];
                    textBox17.Text = In1Config[13].Split('-')[1];
                    textBox18.Text = In1Config[14].Split('-')[1];
                    textBox19.Text = In1Config[15].Split('-')[1];
                    textBox20.Text = In1Config[16].Split('-')[1];
                    textBox21.Text = In1Config[17].Split('-')[1];
                    textBox22.Text = In1Config[18].Split('-')[1];
                    comboBox1.Text = In1Config[19];
                    textBox1.Text = In1Config[20].Substring(0, In1Config[20].Length-3);
                    textBox2.Text = In1Config[21];
                    textBox3.Text = In1Config[22];
                    textBox23.Text = In1Config[23];
                    textBox24.Text = In1Config[24];
                    textBox25.Text = In1Config[25];
                    textBox26.Text = In1Config[26];
                    textBox27.Text = In1Config[27];
                    comboBox6.Text = In1Config[28];
                    comboBox2.Text = In1Config[29];
                    comboBox3.Text = In1Config[30];
                    comboBox7.Text = In1Config[31];
                    comboBox4.Text = In1Config[32];
                    comboBox5.Text = In1Config[33];

                    label2.BackColor = Color.Lime;
                    label3.BackColor = Color.Lime;
                    intervalo = Convert.ToInt32(textBox1.Text +"000");
                    IniciaTimer();
                    MessageBox.Show("Se ha detectado una configuración");
                    try
                    {
                        button2_Click_1(null, null);
                        button4_Click_1(null, null);
                    }
                    catch
                    {
                        MessageBox.Show("Error al autoiniciar el programa");
                        button4.Text = "Iniciar";
                        label2.BackColor = Color.Yellow;
                        label3.BackColor = Color.Yellow;
                        button2.Text = "Conectar";
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("No se ha podido leer el archivo de configuración: \n\n" + ex.ToString());
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        #endregion

        #region Conexion y Carga

        public void CargaPuertos()
        {
            comboBox1.Items.Clear();
            foreach (string Puertos in SerialPort.GetPortNames())
            {
                comboBox1.Items.Add(Puertos);
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            CargaPuertos();
            MessageBox.Show("Los puertos disponibles se han actualizado");
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (button2.Text.Equals("Conectar"))
            {
                if (!comboBox1.Text.Equals("Seleccionar") || !string.IsNullOrEmpty(comboBox1.Text))
                {
                    if (Conectado == false)
                    {
                        serialPort1.PortName = comboBox1.Text;
                        try
                        {
                            serialPort1.Open();
                            button2.Text = "Desconectar";
                            MessageBox.Show("Conectado correctamente");
                            Conectado = true;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("No se pudo establecer la conexión: \n\n" + ex.ToString());
                            label2.BackColor = Color.Yellow;
                            label3.BackColor = Color.Yellow;
                            Conectado = false;
                        }
                    }
                    else
                    {
                        DialogResult dialogResult = MessageBox.Show("Ya está conectado a un puerto serial, desea desconectarse?", "Atención", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {
                            serialPort1.Close();
                            serialPort1.PortName = comboBox1.Text;
                            try
                            {

                                ConfigSerial();
                                serialPort1.Open();
                                button2.Text = "Desconectar";
                                MessageBox.Show("Conectado correctamente");
                                Conectado = true;
                            }
                            catch (Exception r)
                            {
                                MessageBox.Show("No se pudo establecer la conexión: \n\n" + r.ToString());
                                label2.BackColor = Color.Yellow;
                                label3.BackColor = Color.Yellow;
                                Conectado = false;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Seleccione un puerto serial!");
                }
            }
            else if (button2.Text.Equals("Desconectar"))
            {
                serialPort1.Close();
                button2.Text = "Conectar";
                MessageBox.Show("Se ha desconectado del puerto serial");
            }
        }

        #endregion

        #region Selecciona Paths

        private void button5_Click_1(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog1.FileName;
            }
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        #endregion

        #region Inicia y Recibe

        private void button4_Click_1(object sender, EventArgs e)
        {
            if (button4.Text.Equals("Iniciar"))
            {
                RutaArchivo = textBox2.Text;
                RutaCarpeta = textBox3.Text;
                RutaSubcarpeta = Path.Combine(RutaCarpeta, "Copia originales");
                if (Conectado == true)
                {
                    if (!string.IsNullOrEmpty(textBox2.Text) || !string.IsNullOrEmpty(textBox3.Text))
                    {
                        if (!Directory.Exists(RutaSubcarpeta))
                        {
                            Directory.CreateDirectory(RutaSubcarpeta);
                        }
                        try
                        {
                            intervalo = 0;
                            intervalo = Convert.ToInt32(textBox1.Text + "000");
                            try
                            {
                                if(In1Config.Count == 19)
                                {
                                    using (StreamWriter escritor = new StreamWriter(ArchivoConf, true))
                                    {
                                        escritor.WriteLine(comboBox1.Text);
                                        escritor.WriteLine(textBox1.Text);
                                        escritor.WriteLine(textBox2.Text);
                                        escritor.WriteLine(textBox3.Text);
                                        escritor.Close();
                                    }
                                    CargaConfig();
                                }
                                
                                label2.BackColor = Color.Lime;
                                label3.BackColor = Color.Lime;
                                button4.Text = "Detener";
                                IniciaTimer();
                            }
                            catch
                            {
                                MessageBox.Show("Ha ocurrido un error al guardar la configuracíón básica");
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("El intervalo no se ha especificado o se han detectado carácteres inválidos" + ex);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se ha seleccionado una ruta para leer o escribir los archivos");
                    }
                }
                else
                {
                    MessageBox.Show("No se ha detectado una conexión");
                }
            }
            else if(button4.Text.Equals("Detener"))
            {
                aTimer.Enabled = false;
                button4.Text = "Iniciar";
                label2.BackColor = Color.Yellow;
                label3.BackColor = Color.Yellow;
                MessageBox.Show("El proceso se ha detenido");
            }
        }

        public void IniciaTimer()
        {
            aTimer.Interval = intervalo;
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.Enabled = true;

        }

        private void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            RutaArchivo = textBox2.Text;
            RutaCarpeta = textBox3.Text;
            intervalo = Convert.ToInt32(textBox1.Text + "000");
            In1Config[20] = intervalo.ToString();
            In1Config[21] = RutaArchivo;
            In1Config[22] = RutaCarpeta;
            RutaSubcarpeta = Path.Combine(RutaCarpeta, "Copia originales");
            if (!Directory.Exists(RutaSubcarpeta))
            {
                Directory.CreateDirectory(RutaSubcarpeta);
            }
            try
            {
                if (File.Exists(RutaArchivo))
                {
                    try
                    {
                        Fecha = DateTime.Now;
                        File.Copy(RutaArchivo, RutaSubcarpeta + @"\" + Fecha.ToString("yyyyMMdd HHmmss") + ".txt");
                        File.Delete(RutaArchivo);
                        LeeArchivo(RutaSubcarpeta + @"\" + Fecha.ToString("yyyyMMdd HHmmss") + ".txt");
                        label2.BackColor = Color.Lime;
                    }
                    catch (Exception ex)
                    {
                        label2.BackColor = Color.Red;
                    }
                }
            }
            catch (Exception ex)
            {
                label2.BackColor = Color.Red;
            }
        }

        public void LeeArchivo(string Archivo)
        {
            if (File.Exists(Archivo))
            {
                using (StreamReader lector = new StreamReader(Archivo))
                {
                    while ((Linea = lector.ReadLine()) != null)
                    {
                        RecibeCadena(Linea);
                    }
                    lector.Close();
                }
            }
            else
            {
                label2.BackColor = Color.Red;
            }
        }

        #endregion

        #region Procesa Cadena

        public void RecibeCadena(string Cadena)
        {
            In1b = "";
            HOT = "HOT";
            In1m = "";
            In1d = "";
            In1e = "";
            In1h = "";
            In1n = "";
            In1u = "";
            Pesos = "$";
            In1w = "";
            In1v = "";
            In1num = "";
            In1z = "";
            In1t = "";
            In1l = "";
            In1r = "";
            In1i = "";
            In1k = "";
            In1CadenaFinal = "";
            Parte = "";
            In1a = "";
            Partes = new List<string>();

            foreach (string s in Cadena.Split(' '))
            {
                Parte = s.Replace(" ", "");
                if (!string.IsNullOrEmpty(Parte))
                {
                    Partes.Add(Parte);

                }
            }

            try
            {
                In1Aumenta = 0;
                In1Inicia = 0;
                for(int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('b'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('b'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1b = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('m'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('m'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1m = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('d'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('d'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1d = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('e'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('e'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1e = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('h'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('h'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1h = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('n'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('n'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1n = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('u'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('u'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1u = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('w'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('w'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1w = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('v'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('v'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1v = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('#'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('#'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1num = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('z'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('z'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1z = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                if (In1z.Equals("LOC"))
                {
                    In1z = In1Config[23];
                }
                else if (In1z.Equals("CEL"))
                {
                    In1z = In1Config[24];
                }
                else if (In1z.Equals("DDN"))
                {
                    In1z = In1Config[25];
                }
                else if (In1z.Equals("DDI"))
                {
                    In1z = In1Config[26];
                }
                else if (In1z.Equals("TOL"))
                {
                    In1z = In1Config[27];
                }
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('t'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('t'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1t = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('l'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('l'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1l = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('r'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('r'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1r = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('i'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('i'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1i = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('k'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('k'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1k = Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");
                In1Aumenta = 0;
                In1Inicia = 0;
                for (int i = 0; i < In1Config[0].Length; i++)
                {
                    if (In1Config[0][i].Equals('a'))
                    {
                        In1Inicia = i;
                        i = In1Config[0].Length;
                        foreach (char s in In1Config[0])
                        {
                            if (s.Equals('a'))
                            {
                                In1Aumenta++;
                            }
                        }
                    }
                }
                In1a= Cadena.Substring(In1Inicia, In1Aumenta).Replace(" ", "");

                
                for(int i = 0; i < In1Config[1].Length; i++)
                {
                    In1Caracter = In1Config[1][i];
                    In1Inicia = i;
                    In1Aumenta = 0;
                    foreach (char s in In1Config[1])
                    {
                        if (s.Equals(In1Caracter))
                        {
                            In1Aumenta++;
                        }
                    }
                    if (In1Caracter.Equals('a'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[2].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[2].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1a.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[2].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1a;
                            }
                            else if (In1Config[2].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1a;
                                for (int n = 0; n < (In1Aumenta - In1a.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[2].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1a;
                        }
                    }
                    else if (In1Caracter.Equals('m'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[3].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[3].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1m.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[3].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1m;
                            }
                            else if (In1Config[3].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1m;
                                for (int n = 0; n < (In1Aumenta - In1m.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[3].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1m;
                        }
                    }
                    else if (In1Caracter.Equals('d'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[4].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[4].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1d.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[4].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1d;
                            }
                            else if (In1Config[4].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1d;
                                for (int n = 0; n < (In1Aumenta - In1d.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[4].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1d;
                        }
                    }
                    else if (In1Caracter.Equals('h'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[5].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[5].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1h.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[5].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1h;
                            }
                            else if (In1Config[5].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1h;
                                for (int n = 0; n < (In1Aumenta - In1h.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[5].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1h;
                        }
                    }
                    else if (In1Caracter.Equals('n'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[6].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[6].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1n.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[6].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1n;
                            }
                            else if (In1Config[6].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1n;
                                for (int n = 0; n < (In1Aumenta - In1n.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[6].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1n;
                        }
                    }
                    else if (In1Caracter.Equals('e'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[7].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[7].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1e.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[7].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1e;
                            }
                            else if (In1Config[7].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1e;
                                for (int n = 0; n < (In1Aumenta - In1e.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[7].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1e;
                        }
                    }
                    else if (In1Caracter.Equals('b'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[8].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[8].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1b.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[8].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1b;
                            }
                            else if (In1Config[8].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1b;
                                for (int n = 0; n < (In1Aumenta - In1b.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[8].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1b;
                        }
                    }
                    else if (In1Caracter.Equals('t'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[9].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[9].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1t.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[9].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1t;
                            }
                            else if (In1Config[9].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1t;
                                for (int n = 0; n < (In1Aumenta - In1t.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[9].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1t;
                        }
                    }
                    else if (In1Caracter.Equals('#'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[10].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[10].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1num.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[10].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1num;
                            }
                            else if (In1Config[10].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1num;
                                for (int n = 0; n < (In1Aumenta - In1num.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[10].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1num;
                        }
                    }
                    else if (In1Caracter.Equals('u'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[11].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[11].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1u.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[11].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1u;
                            }
                            else if (In1Config[11].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1u;
                                for (int n = 0; n < (In1Aumenta - In1u.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[11].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1u;
                        }
                    }
                    else if (In1Caracter.Equals('l'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[12].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[12].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1l.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[12].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1l;
                            }
                            else if (In1Config[12].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1l;
                                for (int n = 0; n < (In1Aumenta - In1l.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[12].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1l;
                        }
                    }
                    else if (In1Caracter.Equals('r'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[13].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[13].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1r.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[13].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1r;
                            }
                            else if (In1Config[13].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1r;
                                for (int n = 0; n < (In1Aumenta - In1r.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[13].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1r;
                        }
                    }
                    else if (In1Caracter.Equals('i'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[14].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[14].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1i.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[14].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1i;
                            }
                            else if (In1Config[14].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1i;
                                for (int n = 0; n < (In1Aumenta - In1i.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[14].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1i;
                        }
                    }
                    else if (In1Caracter.Equals('w'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[15].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[15].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1w.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[15].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1w;
                            }
                            else if (In1Config[15].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1w;
                                for (int n = 0; n < (In1Aumenta - In1w.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[15].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1w;
                        }
                    }
                    else if (In1Caracter.Equals('v'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[16].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[16].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1v.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[16].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1v;
                            }
                            else if (In1Config[16].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1v;
                                for (int n = 0; n < (In1Aumenta - In1v.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[16].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1v;
                        }
                    }
                    else if (In1Caracter.Equals('k'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[17].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[17].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1k.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[17].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1k;
                            }
                            else if (In1Config[17].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1k;
                                for (int n = 0; n < (In1Aumenta - In1k.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[17].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1k;
                        }
                    }
                    else if (In1Caracter.Equals('z'))
                    {
                        i = In1Inicia + In1Aumenta;
                        i--;
                        if (!In1Config[18].Split('-')[1].Equals(" "))
                        {
                            if (In1Config[18].Split('-')[1][1].Equals('d'))
                            {
                                for (int n = 0; n < (In1Aumenta - In1z.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[18].Split('-')[1][0];
                                }
                                In1CadenaFinal += In1z;
                            }
                            else if (In1Config[18].Split('-')[1][1].Equals('i'))
                            {
                                In1CadenaFinal += In1z;
                                for (int n = 0; n < (In1Aumenta - In1z.Length); n++)
                                {
                                    In1CadenaFinal += In1Config[18].Split('-')[1][0];
                                }
                            }
                        }
                        else
                        {
                            In1CadenaFinal += In1z;
                        }
                    }
                    else
                    {
                        In1CadenaFinal += In1Caracter;
                    }
                }

                //Envia(CadenaFinal);
                MessageBox.Show(In1CadenaFinal);
                listBox1.Invoke(new Action(() => { listBox1.Items.Add(In1CadenaFinal); }));

            }
            catch (Exception ex)
            {
                listBox1.Invoke(new Action(() => { listBox1.Items.Add("Error en la siguiente cadena: " + Cadena); }));
                label3.BackColor = Color.Red;
                GuardaErroneas(Cadena);
            }
        }


        #endregion

        #region Envia y guarda

        byte[] ba;
        int value;
        string stringValue;
        string hexString;
        public void Envia(string Cadena)
        {
            try
            {
                ConfigSerial();
                ba = Encoding.Default.GetBytes(Cadena);
                hexString = BitConverter.ToString(ba);
                hexString = hexString.Replace("-", " ");
                hexString = hexString + " 0D 0A";
                Cadena = "";
                foreach (string hex in hexString.Split(' '))
                {
                    value = Convert.ToInt32(hex, 16);
                    stringValue = Char.ConvertFromUtf32(value);
                    Cadena = Cadena + stringValue;
                }
                serialPort1.Write(Cadena);
                GuardaEnviadas(Cadena);
                
            }
            catch (Exception ex)
            {
                label3.BackColor = Color.Red;
                MessageBox.Show(ex.ToString());
            }
        }

        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //MessageBox.Show("|" + serialPort1.ReadExisting() + "|");
        }

        public void GuardaEnviadas(string Cadena)
        {
            SubCarpetaEnviados = Path.Combine(RutaCarpeta, "Enviados");
            if (!Directory.Exists(SubCarpetaEnviados))
            {
                Directory.CreateDirectory(SubCarpetaEnviados);
            }
            ArchivoEnviados = Path.Combine(SubCarpetaEnviados, "Enviados.txt");
            using (StreamWriter escritor = new StreamWriter(ArchivoEnviados, true))
            {
                escritor.WriteLine(Cadena);
                escritor.Close();
            }
        }
        public void GuardaErroneas(string Cadena)
        {
            try
            {
                SubCarpetaErroneos = Path.Combine(RutaCarpeta, "Erroneos");
                if (!Directory.Exists(SubCarpetaErroneos))
                {
                    Directory.CreateDirectory(SubCarpetaErroneos);
                }
                ArchivoErroneos = Path.Combine(SubCarpetaErroneos, "Erroneos.txt");
                using (StreamWriter escritor = new StreamWriter(ArchivoErroneos, true))
                {
                    escritor.WriteLine(Cadena);
                    escritor.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public void GuardaRechazados(string Cadena)
        {
            SubCarpetaEnviados = Path.Combine(RutaCarpeta, "Rechazados");
            if (!Directory.Exists(SubCarpetaEnviados))
            {
                Directory.CreateDirectory(SubCarpetaEnviados);
            }
            ArchivoEnviados = Path.Combine(SubCarpetaEnviados, "Rechazados.txt");
            using (StreamWriter escritor = new StreamWriter(ArchivoEnviados, true))
            {
                escritor.WriteLine(Cadena);
                escritor.Close();
            }
        }

        #endregion

        #region Parámetros

        
        public void ConfigSerial()
        {
            serialPort1.BaudRate = Convert.ToInt32(In1Config[28]);

            if (In1Config[29].Equals("EVEN")) { serialPort1.Parity = Parity.Even; }
            else if (In1Config[29].Equals("MARK")) { serialPort1.Parity = Parity.Mark; }
            else if (In1Config[29].Equals("ODD")) { serialPort1.Parity = Parity.Odd; }
            else if (In1Config[29].Equals("NONE")) { serialPort1.Parity = Parity.None; }
            else if (In1Config[29].Equals("SPACE")) { serialPort1.Parity = Parity.Space; }

            if (In1Config[30].Equals("NONE")) { serialPort1.StopBits = StopBits.None; }
            else if (In1Config[30].Equals("ONE")) { serialPort1.StopBits = StopBits.One; }
            else if (In1Config[30].Equals("ONE POINT FIVE")) { serialPort1.StopBits = StopBits.OnePointFive; }
            else if (In1Config[30].Equals("TWO")) { serialPort1.StopBits = StopBits.Two; }

            serialPort1.DataBits = Convert.ToInt32(In1Config[31]);

            if (In1Config[32].Equals("NONE")) { serialPort1.Handshake = Handshake.None; }
            else if (In1Config[32].Equals("REQUEST TO SEND")) { serialPort1.Handshake = Handshake.RequestToSend; }

            if (In1Config[33].Equals("TRUE")) { serialPort1.RtsEnable = true; }
            else if (In1Config[33].Equals("FALSE")) { serialPort1.RtsEnable = false; }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    using (StreamReader lector = new StreamReader(openFileDialog1.FileName))
                    {
                        label45.Text = lector.ReadLine();
                    }
                }
            }
            catch (Exception ex)
            {
                label45.Text = "";
                MessageBox.Show(ex.ToString());
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            label46.Text = "";
            for (int i = 0; i < textBox4.Text.Length - 1; i++)
            {
                label46.Text += " ";
            }
            label46.Text += "|";
        }

        private void label45_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(label45.Text))
            {
                textBox4.Enabled = false;
            }
            else
            {
                textBox4.Enabled = true;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    using (StreamReader lector = new StreamReader(openFileDialog1.FileName))
                    {
                        label48.Text = lector.ReadLine();
                    }
                }
            }
            catch (Exception ex)
            {
                label48.Text = "";
                MessageBox.Show(ex.ToString());
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            label47.Text = "";
            for (int i = 0; i < textBox5.Text.Length - 1; i++)
            {
                label47.Text += " ";
            }
            label47.Text += "|";
        }

        private void label48_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(label48.Text))
            {
                textBox5.Enabled = false;
            }
            else
            {
                textBox5.Enabled = true;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists(ArchivoConf))
                {
                    File.Delete(ArchivoConf);
                }
                using (StreamWriter escritor = new StreamWriter(ArchivoConf))
                {
                    escritor.WriteLine(textBox4.Text);
                    escritor.WriteLine(textBox5.Text);
                    if (string.IsNullOrEmpty(textBox6.Text)) { escritor.WriteLine("a- "); } else { escritor.WriteLine("a-" + textBox6.Text); }
                    if (string.IsNullOrEmpty(textBox7.Text)) { escritor.WriteLine("m- "); } else { escritor.WriteLine("m-" + textBox7.Text); }
                    if (string.IsNullOrEmpty(textBox8.Text)) { escritor.WriteLine("d- "); } else { escritor.WriteLine("d-" + textBox8.Text); }
                    if (string.IsNullOrEmpty(textBox9.Text)) { escritor.WriteLine("h- "); } else { escritor.WriteLine("h-" + textBox9.Text); }
                    if (string.IsNullOrEmpty(textBox10.Text)) { escritor.WriteLine("n- "); } else { escritor.WriteLine("n-" + textBox10.Text); }
                    if (string.IsNullOrEmpty(textBox11.Text)) { escritor.WriteLine("e- "); } else { escritor.WriteLine("e-" + textBox11.Text); }
                    if (string.IsNullOrEmpty(textBox12.Text)) { escritor.WriteLine("b- "); } else { escritor.WriteLine("b-" + textBox12.Text); }
                    if (string.IsNullOrEmpty(textBox13.Text)) { escritor.WriteLine("t- "); } else { escritor.WriteLine("t-" + textBox13.Text); }
                    if (string.IsNullOrEmpty(textBox14.Text)) { escritor.WriteLine("#- "); } else { escritor.WriteLine("#-" + textBox14.Text); }
                    if (string.IsNullOrEmpty(textBox15.Text)) { escritor.WriteLine("u- "); } else { escritor.WriteLine("u-" + textBox15.Text); }
                    if (string.IsNullOrEmpty(textBox16.Text)) { escritor.WriteLine("l- "); } else { escritor.WriteLine("l-" + textBox16.Text); }
                    if (string.IsNullOrEmpty(textBox17.Text)) { escritor.WriteLine("r- "); } else { escritor.WriteLine("r-" + textBox17.Text); }
                    if (string.IsNullOrEmpty(textBox18.Text)) { escritor.WriteLine("i- "); } else { escritor.WriteLine("i-" + textBox18.Text); }
                    if (string.IsNullOrEmpty(textBox19.Text)) { escritor.WriteLine("w- "); } else { escritor.WriteLine("w-" + textBox19.Text); }
                    if (string.IsNullOrEmpty(textBox20.Text)) { escritor.WriteLine("v- "); } else { escritor.WriteLine("v-" + textBox20.Text); }
                    if (string.IsNullOrEmpty(textBox21.Text)) { escritor.WriteLine("k- "); } else { escritor.WriteLine("k-" + textBox21.Text); }
                    if (string.IsNullOrEmpty(textBox22.Text)) { escritor.WriteLine("z- "); } else { escritor.WriteLine("z-" + textBox22.Text); }
                    escritor.WriteLine(comboBox1.Text);
                    escritor.WriteLine(textBox1.Text + "000");
                    escritor.WriteLine(textBox2.Text);
                    escritor.WriteLine(textBox3.Text);
                    escritor.WriteLine(textBox23.Text);
                    escritor.WriteLine(textBox24.Text);
                    escritor.WriteLine(textBox25.Text);
                    escritor.WriteLine(textBox26.Text);
                    escritor.WriteLine(textBox27.Text);
                    escritor.WriteLine(comboBox6.Text);
                    escritor.WriteLine(comboBox2.Text);
                    escritor.WriteLine(comboBox3.Text);
                    escritor.WriteLine(comboBox7.Text);
                    escritor.WriteLine(comboBox4.Text);
                    escritor.WriteLine(comboBox5.Text);
                    escritor.Close();
                }
                if(apagando == false)
                {
                    DialogResult dialogResult = MessageBox.Show("Para efectuar los cambios el programa deberá reiniciarse, ¿Desea continuar?", "Atención", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        Application.Restart();
                    }
                    else
                    {
                        MessageBox.Show("Los cambios no se aplicarán hasta que se reinicie el programa");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se ha podido guardar el archivo: \n\n" + ex.ToString());
            }
        }


        #endregion

        private void label45_Click(object sender, EventArgs e)
        {

        }
    }
}
