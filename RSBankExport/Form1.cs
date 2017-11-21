using System;
using System.IO;
using System.Windows.Forms;
using System.Xml.Linq;

namespace RSBankExport
{
    public partial class Form1 : Form
    {
        static int EMPLOYERS_COUNT = 0;
        static int CUR_EXPORT_NUMBER = 0;
        static String fullPath = Application.StartupPath.ToString();
        static employer[] EMPLOYERS_LIST = new employer[EMPLOYERS_COUNT];

        public Form1()
        {
            InitializeComponent();
            onLoad();
            label6.Text = "Прочитано " + EMPLOYERS_COUNT + " сотрудник{ов/а}, имеющих карты РусЮгБанка" + "\n";
            textBox1.Text = Convert.ToString(CUR_EXPORT_NUMBER);
        }

        public static void onLoad()
        {
            System.IO.DirectoryInfo dirSource = new System.IO.DirectoryInfo(fullPath + "\\Export");
            if (!dirSource.Exists)
            {
                System.IO.Directory.CreateDirectory(fullPath + "\\Export");
            }
            System.IO.FileStream fCh = new System.IO.FileStream(fullPath + "\\Export\\12253_00.txt", System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Read);
            System.IO.StreamReader SR = new System.IO.StreamReader(fCh, System.Text.Encoding.Default);

            string str = "";
            while (!SR.EndOfStream)
            {
                str = SR.ReadLine();
                if (str == "")
                    continue;
                EMPLOYERS_COUNT++;
            }
            SR.Close();
            fCh.Close();

            Array.Resize<employer>(ref EMPLOYERS_LIST, EMPLOYERS_COUNT);
            
            fCh = new System.IO.FileStream(fullPath + "\\Export\\12253_00.txt", System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Read);
            SR = new System.IO.StreamReader(fCh, System.Text.Encoding.Default);
            int i = 0;
            while (!SR.EndOfStream)
            {
                str = SR.ReadLine();
                if (str == "")
                    continue;
                String[] words = str.Split(new char[] { '^' }, StringSplitOptions.RemoveEmptyEntries);
                EMPLOYERS_LIST[i] = new employer(words[2], words[3], words[4], 0.0);
                EMPLOYERS_LIST[i].account = words[0];
                EMPLOYERS_LIST[i].bik = words[1];
                EMPLOYERS_LIST[i].series = words[5];
                EMPLOYERS_LIST[i].code = words[6];
                i++;
            }
            SR.Close();
            fCh.Close();
            fCh = new System.IO.FileStream(fullPath + "\\RSBank.dat", System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite);
            System.IO.StreamReader cur_num_stream = new System.IO.StreamReader(fCh);
            CUR_EXPORT_NUMBER = Convert.ToInt32(cur_num_stream.ReadLine()) + 1;
            cur_num_stream.Close();
            fCh.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int loaded = 0;
            double sum = 0.0;
            System.Windows.Forms.OpenFileDialog dlg = new System.Windows.Forms.OpenFileDialog();
            dlg.Filter = "Файлы XML (*.xml)|*.xml";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                XDocument xdoc = null;
                try
                {
                    xdoc = XDocument.Load(dlg.FileName);

                    int i = 0;
                    label_DateExport.Text = xdoc.Root.Attribute("ДатаФормирования").Value;
                    foreach (XElement xe in xdoc.Element("СчетаПК").Elements("ЗачислениеЗарплаты").Elements("Сотрудник"))
                    {
                        for (int j = 0; j < EMPLOYERS_COUNT; j++) // j = i, заменено ввиду представления списка не по алфавиту
                        {
                            if (xe.Element("Фамилия").Value.ToUpper() == EMPLOYERS_LIST[j].lName &&
                                xe.Element("Имя").Value.ToUpper() == EMPLOYERS_LIST[j].fName &&
                                xe.Element("Отчество").Value.ToUpper() == EMPLOYERS_LIST[j].mName)
                            {
                                EMPLOYERS_LIST[j].money = double.Parse(xe.Element("Сумма").Value, System.Globalization.CultureInfo.InvariantCulture);
                                sum += EMPLOYERS_LIST[j].money;
                                i = j;
                                loaded++;
                                break;
                            }
                        }
                    }
                    label6.Text += "Обработано " + loaded + " сотрудник{ов/а} из файла 1С" + "\n";
                    label_Sum.Text = sum.ToString("0.00", System.Globalization.CultureInfo.GetCultureInfo("en-US"));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка: " + ex.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CUR_EXPORT_NUMBER = Convert.ToInt32(textBox1.Text);
            string fileName = EMPLOYERS_LIST[0].code + "_";
            if (CUR_EXPORT_NUMBER < 10)
                fileName += "0";
            fileName += CUR_EXPORT_NUMBER + ".txt";

            FileStream fs = new FileStream(fullPath + "\\Export\\" + fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter SW = new StreamWriter(fs);
            for (int i = 0; i < EMPLOYERS_COUNT; i++)
            {
                if (EMPLOYERS_LIST[i].money != 0)
                {
                    string line = EMPLOYERS_LIST[i].account + "^" +
                                 EMPLOYERS_LIST[i].bik + "^" +
                                 EMPLOYERS_LIST[i].money.ToString("0.00", System.Globalization.CultureInfo.GetCultureInfo("en-US")) + "^" +
                                 EMPLOYERS_LIST[i].lName + "^" +
                                 EMPLOYERS_LIST[i].fName + "^" +
                                 EMPLOYERS_LIST[i].mName + "^" +
                                 EMPLOYERS_LIST[i].series + "^" +
                                 EMPLOYERS_LIST[i].code + "^";
                    SW.WriteLine(line);
                }
            }
            SW.Close();
            fs.Close();
            label6.Text += "Файл сохранён с именем: " + fileName + "\n"
                        + "Можно завершить работу программы." + "\n";
            FileStream fCh = new FileStream(fullPath + "\\RSBank.dat", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            SW = new StreamWriter(fCh);
            SW.WriteLine(textBox1.Text);
            SW.Close();
            fCh.Close();
        }
    }
    public class employer
    {
        public string fName;
        public string lName;
        public string mName;
        public string account;
        public string bik;
        public string series;
        public string code;

        public double money;

        public employer(string _lName, String _fName, String _mName, Double _money)
        {
            this.fName = _fName;
            this.lName = _lName;
            this.mName = _mName;
            this.money = _money;
        }
    }
}