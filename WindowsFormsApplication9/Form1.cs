using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.Xml;
using Graph = Microsoft.Office.Interop.Graph;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Threading;
using gma.System.Windows;
namespace WindowsFormsApplication9
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            // backgroundWorker1.RunWorkerAsync();

        }
        List<question> questions = new List<question>();
        create_new_pres player = new create_new_pres();
        Color[] bar_colors = new Color[8];

        List<Thread> oThread = new List<Thread>();
        List<Thread> rThread = new List<Thread>();
        int data_format = 2;
        public int question_no = 0;
        int choices_no1 = 0;
        int num = 1;
        int rr = 0;
        
        string fastest;
        public Random random = new Random();

        List<Team> teams = new List<Team>(30);
      List<Team> teams1 = new List<Team>(3);
        

        public void addd_ata()
        {


            //  XmlTextReader reader1 = new XmlTextReader(folderBrowserDialog2.SelectedPath + "\\" + textBox2.Text + ".xml");
            //XmlTextReader reader = new XmlTextReader(folderBrowserDialog1.SelectedPath+"\\"+textBox1.Text+".xml");
            //XmlTextReader reader = new XmlTextReader("\\\\voting-1\\test-1\\shendy1.xml");
            //XmlTextReader reader1 = new XmlTextReader("\\\\voting-1\\test-1\\shendy1.xml");
            XmlTextReader reader = new XmlTextReader("c:\\users\\ali\\desktop\\whats your name.xml");
            XmlTextReader reader1 = new XmlTextReader("c:\\users\\ali\\desktop\\whats your name.xml");
            

            for (int i = 0; i <= question_no; i++)
            {
                reader.ReadToFollowing("Question");
            }
            if (!reader.EOF)
            {

                reader.MoveToAttribute(1);
                questions[question_no].title = reader.Value;

                reader.MoveToAttribute(5);
                questions[question_no].choices_no = reader.Value;
                choices_no1 = Convert.ToInt32(questions[question_no].choices_no);
                //reader.ReadToFollowing("Data"); 
                for (int i = 0; i < choices_no1; ++i)
                {

                    reader.ReadToFollowing("Data");
                    reader.MoveToFirstAttribute();
                    questions[question_no].choices.Add(reader.Value);
                    reader.MoveToContent();

                    questions[question_no].values.Add(reader.ReadElementContentAsInt());

                }




            }
            for (int i = 0; i <= question_no; i++)
            {
                reader1.ReadToFollowing("Question");
            }
            if (!reader1.EOF)
            {

                // reader1.ReadToFollowing("Data");

                for (int i = 0; i < choices_no1; i++)
                {
                    reader1.ReadToFollowing("Data");

                    reader1.MoveToContent();

                    questions[question_no].values[i] = (questions[question_no].values[i] + reader1.ReadElementContentAsInt());

                }









                player.add_slide(ref question_no, ref questions, ref openFileDialog1);
                player.add_title(ref questions, ref question_no);
                player.add_choices(ref questions, ref question_no, ref choices_no1);
                player.add_chart(ref questions, ref question_no, ref choices_no1, ref colorDialog1, ref data_format, ref bar_colors, ref fontDialog1);
                player.start_show(ref question_no);
                question_no++;



            }


            reader.Close();
            reader1.Close();
            player.objApp.ActivePresentation.Save();


        }













        private void titlef_Click(object sender, EventArgs e)
        {
            DialogResult dResult = fontDialog2.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                questions[comboBox1.SelectedIndex].title_font = fontDialog2.Font;

            }
        }

        private void titlec_Click(object sender, EventArgs e)
        {
            DialogResult dResult = colorDialog12.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                pictureBox11.BackColor = colorDialog12.Color;
                questions[comboBox1.SelectedIndex].Title_color = colorDialog12.Color;
            }
        }

        private void backpicb_Click(object sender, EventArgs e)
        {
            DialogResult dResult = colorDialog1.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                backpic.BackColor = colorDialog1.Color;
                questions[comboBox1.SelectedIndex].background = colorDialog1.Color;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            folderBrowserDialog1.SelectedPath = Application.StartupPath;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            folderBrowserDialog2.ShowDialog();
        }

        private void b1_Click(object sender, EventArgs e)
        {
            DialogResult dResult = colorDialog2.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                bar1.BackColor = colorDialog2.Color;
                bar_colors[0] = colorDialog2.Color;

            }

        }

        private void b2_Click(object sender, EventArgs e)
        {

            DialogResult dResult = colorDialog3.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                bar2.BackColor = colorDialog3.Color;
                bar_colors[1] = colorDialog3.Color;

            }
        }

        private void b3_Click(object sender, EventArgs e)
        {
            DialogResult dResult = colorDialog4.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                bar3.BackColor = colorDialog4.Color;
                bar_colors[2] = colorDialog4.Color;

            }
        }

        private void b4_Click(object sender, EventArgs e)
        {
            DialogResult dResult = colorDialog5.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                bar4.BackColor = colorDialog5.Color;
                bar_colors[3] = colorDialog5.Color;

            }
        }

        private void b5_Click(object sender, EventArgs e)
        {
            DialogResult dResult = colorDialog6.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                bar5.BackColor = colorDialog6.Color;
                bar_colors[4] = colorDialog6.Color;

            }
        }

        private void b6_Click(object sender, EventArgs e)
        {
            DialogResult dResult = colorDialog7.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                bar6.BackColor = colorDialog7.Color;
                bar_colors[5] = colorDialog7.Color;

            }
        }

        private void b7_Click(object sender, EventArgs e)
        {
            DialogResult dResult = colorDialog8.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                bar7.BackColor = colorDialog8.Color;
                bar_colors[6] = colorDialog8.Color;

            }
        }

        private void b8_Click(object sender, EventArgs e)
        {
            DialogResult dResult = colorDialog9.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                bar8.BackColor = colorDialog9.Color;
                bar_colors[7] = colorDialog9.Color;
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {



            if (num > Convert.ToInt32(numericUpDown1.Value))
            {
                for (int i = num; i > Convert.ToInt32(numericUpDown1.Value); i--)
                {
                    questions.RemoveAt(i - 1);

                    comboBox1.Items.RemoveAt(i - 1);
                }

                comboBox1.MaxDropDownItems = Convert.ToInt32(numericUpDown1.Value);
                num = Convert.ToInt32(numericUpDown1.Value);
            }
            else if (num < Convert.ToInt32(numericUpDown1.Value))
            {
                comboBox1.MaxDropDownItems = Convert.ToInt32(numericUpDown1.Value);
                questions.Capacity = comboBox1.MaxDropDownItems;
                for (int i = num; i < Convert.ToInt32(numericUpDown1.Value); i++)
                {

                    questions.Insert(i, new question());
                    questions[i].Title_color = colorDialog12.Color;
                    questions[i].title_font = fontDialog2.Font;
                    questions[i].choices_color = colorDialog14.Color;
                    questions[i].choices_font = fontDialog3.Font;
                    questions[i].show_chart = checkBox4.Checked;
                    questions[i].show_choices = checkBox3.Checked;
                    questions[i].show_title = checkBox2.Checked;
                    questions[i].background = colorDialog1.Color;
                    comboBox1.Items.Insert(i, "Question" + (i + 1).ToString());
                    questions[i].orien = true;
                }

                num = Convert.ToInt32(numericUpDown1.Value);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            pictureBox10.BackColor = questions[comboBox1.SelectedIndex].choices_color;
            pictureBox11.BackColor = questions[comboBox1.SelectedIndex].Title_color;
            backpic.BackColor = questions[comboBox1.SelectedIndex].background;




            checkBox4.Checked = questions[comboBox1.SelectedIndex].show_chart;
            checkBox3.Checked = questions[comboBox1.SelectedIndex].show_choices;
            checkBox2.Checked = questions[comboBox1.SelectedIndex].show_title;
            if (questions[comboBox1.SelectedIndex].orien == true)
                radioButton5.Checked = true;
            else
                radioButton4.Checked = true;
            openFileDialog1.FileName = questions[comboBox1.SelectedIndex].back_image;

        }

        private void choicef_Click(object sender, EventArgs e)
        {

            DialogResult dResult = fontDialog3.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                questions[comboBox1.SelectedIndex].choices_font = fontDialog3.Font;


            }
        }

        private void choicec_Click(object sender, EventArgs e)
        {
            DialogResult dResult = colorDialog14.ShowDialog();
            if (dResult == DialogResult.OK)
            {
                pictureBox10.BackColor = colorDialog14.Color;
                questions[comboBox1.SelectedIndex].choices_color = colorDialog14.Color;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                data_format = 1;

            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                data_format = 2;

            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked == true)
            {
                data_format = 3;

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (button5.Enabled == false)
                button5.Enabled = true;
            else
                button5.Enabled = false;
        }
        UserActivityHook actHook;
        UserActivityHook actHook1;
        private void Form1_Load(object sender, EventArgs e)
        {


            int num = 1;
            comboBox1.MaxDropDownItems = num;
            comboBox1.Items.Insert(0, "Question1");

            questions.Capacity = comboBox1.MaxDropDownItems;
            questions.Insert(0, new question());
            num = Convert.ToInt32(numericUpDown1.Value);
            questions[0].Title_color = colorDialog12.Color;
            questions[0].title_font = fontDialog2.Font;
            questions[0].choices_color = colorDialog14.Color;
            questions[0].choices_font = fontDialog3.Font;
            questions[0].show_chart = checkBox4.Checked;
            questions[0].show_choices = checkBox3.Checked;
            questions[0].show_title = checkBox2.Checked;
            questions[0].background = colorDialog1.Color;
            questions[0].orien = true;
            comboBox1.SelectedIndex = 0;




            player.Init();

            bar_colors[0] = colorDialog2.Color;
            bar_colors[1] = colorDialog3.Color;
            bar_colors[2] = colorDialog4.Color;
            bar_colors[3] = colorDialog5.Color;
            bar_colors[4] = colorDialog6.Color;
            bar_colors[5] = colorDialog7.Color;
            bar_colors[6] = colorDialog8.Color;
            bar_colors[7] = colorDialog9.Color;

            bar1.BackColor = colorDialog2.Color;
            bar2.BackColor = colorDialog3.Color;
            bar3.BackColor = colorDialog4.Color;
            bar4.BackColor = colorDialog5.Color;
            bar5.BackColor = colorDialog6.Color;
            bar6.BackColor = colorDialog7.Color;
            bar7.BackColor = colorDialog8.Color;
            bar8.BackColor = colorDialog9.Color;

            pictureBox10.BackColor = questions[comboBox1.SelectedIndex].choices_color;
            pictureBox11.BackColor = questions[comboBox1.SelectedIndex].Title_color;
            backpic.BackColor = questions[comboBox1.SelectedIndex].background;

        }




        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {













        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //  XmlTextReader reader1 = new XmlTextReader(folderBrowserDialog2.SelectedPath + "\\" + textBox2.Text + ".xml");
            //XmlTextReader reader = new XmlTextReader(folderBrowserDialog1.SelectedPath+"\\"+textBox1.Text+".xml");
            //XmlTextReader reader = new XmlTextReader("\\\\voting-1\\test-1\\shendy1.xml");
            //XmlTextReader reader1 = new XmlTextReader("\\\\voting-1\\test-1\\shendy1.xml");
            XmlTextReader reader = new XmlTextReader("c:\\users\\ali\\desktop\\whats your name.xml");
            XmlTextReader reader1 = new XmlTextReader("c:\\users\\ali\\desktop\\whats your name.xml");

            if (player.flag == 1)
            {
                player.Init();
                question_no = 0;
            }
            for (int i = 0; i <= question_no; i++)
            {
                reader.ReadToFollowing("Question");
            }
            if (!reader.EOF)
            {

                reader.MoveToAttribute(1);
                questions[question_no].title = reader.Value;

                reader.MoveToAttribute(5);
                questions[question_no].choices_no = reader.Value;
                choices_no1 = Convert.ToInt32(questions[question_no].choices_no);
                //reader.ReadToFollowing("Data"); 
                for (int i = 0; i < choices_no1; ++i)
                {

                    reader.ReadToFollowing("Data");
                    reader.MoveToFirstAttribute();
                    questions[question_no].choices.Add(reader.Value);
                    reader.MoveToContent();

                    questions[question_no].values.Add(reader.ReadElementContentAsInt());

                }




            }
            for (int i = 0; i <= question_no; i++)
            {
                reader1.ReadToFollowing("Question");
            }
            if (!reader1.EOF)
            {

                // reader1.ReadToFollowing("Data");

                for (int i = 0; i < choices_no1; i++)
                {
                    reader1.ReadToFollowing("Data");

                    reader1.MoveToContent();

                    questions[question_no].values[i] = (questions[question_no].values[i] + reader1.ReadElementContentAsInt());

                }









                player.add_slide(ref question_no, ref questions, ref openFileDialog1);
                player.add_title(ref questions, ref question_no);
                player.add_choices(ref questions, ref question_no, ref choices_no1);
                player.add_chart(ref questions, ref question_no, ref choices_no1, ref colorDialog1, ref data_format, ref bar_colors, ref fontDialog1);
                player.start_show(ref question_no);
                question_no++;



            }


            reader.Close();
            reader1.Close();
            actHook = new UserActivityHook(false, true);
            actHook.KeyPress += new KeyPressEventHandler(key_pressed);
            button1.Text = "Show Started";
            button1.Enabled = false;
            player.objApp.ActivePresentation.SaveAs(saveFileDialog1.FileName);

        }
        private void key_pressed(object sender, KeyPressEventArgs e)
        {
            if (question_no < numericUpDown1.Value)
            {
                if (e.KeyChar == 'd')
                {



                    oThread.Add(new Thread(new ThreadStart(addd_ata)));

                    oThread[rr].Start();
                    if (rr != 0)
                        oThread[rr - 1].Abort();



                    rr++;


                }



            }

            if (e.KeyChar == 'g')
            {
                try
                {
                    oThread.Add(new Thread(new ThreadStart(revote)));
                    oThread[rr].Start();
                    if (rr != 0)
                        oThread[rr - 1].Abort();
                    rr++;
                }
                catch (Exception es)
                {
                    button1.Text = es.Message;
                }
            }
            

        }
        private void key_pressed1(object sender, KeyPressEventArgs e)
        {
           
                if (e.KeyChar == 'k')
                {



                    oThread.Add(new Thread(new ThreadStart(updt1)));

                    oThread[rr].Start();
                    if (rr != 0)
                        oThread[rr - 1].Abort();



                    rr++;


                }



            

            //if (e.KeyChar == 'g')
            //{
            //    try
            //    {
            //        oThread.Add(new Thread(new ThreadStart(revote)));
            //        oThread[rr].Start();
            //        if (rr != 0)
            //            oThread[rr - 1].Abort();
            //        rr++;
            //    }
            //    catch (Exception es)
            //    {
            //        button1.Text = es.Message;
            //    }
            //}
        }
        private void Form1_Validated(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            questions[comboBox1.SelectedIndex].show_title = checkBox2.Checked;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            questions[comboBox1.SelectedIndex].show_choices = checkBox3.Checked;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            questions[comboBox1.SelectedIndex].show_chart = checkBox4.Checked;
        }

        private void fontDialog3_Apply(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {

        }
        private void revote()
        {
            //  XmlTextReader reader1 = new XmlTextReader(folderBrowserDialog2.SelectedPath + "\\" + textBox2.Text + ".xml");
            //XmlTextReader reader = new XmlTextReader(folderBrowserDialog1.SelectedPath+"\\"+textBox1.Text+".xml");
            //XmlTextReader reader = new XmlTextReader("\\\\voting-1\\test-1\\shendy1.xml");
            //XmlTextReader reader1 = new XmlTextReader("\\\\voting-1\\test-1\\shendy1.xml");
            XmlTextReader reader = new XmlTextReader("c:\\users\\ali\\desktop\\whats your name.xml");
            XmlTextReader reader1 = new XmlTextReader("c:\\users\\ali\\desktop\\whats your name.xml");
            int s = player.objApp.ActivePresentation.SlideShowWindow.View.Slide.SlideIndex - 1;
            if (player.flag == 1)
            {
                player.Init();
                question_no = 0;
            }
            for (int i = 0; i <= s; i++)
            {
                reader.ReadToFollowing("Question");
            }
            if (!reader.EOF)
            {




                choices_no1 = Convert.ToInt32(questions[s].choices_no);
                //reader.ReadToFollowing("Data"); 
                for (int i = 0; i < choices_no1; ++i)
                {
                    reader.ReadToFollowing("Data");

                    reader.MoveToContent();

                    questions[s].values[i] = (reader.ReadElementContentAsInt());

                }




            }
            for (int i = 0; i <= s; i++)
            {
                reader1.ReadToFollowing("Question");
            }
            if (!reader1.EOF)
            {

                //reader1.ReadToFollowing("Data");

                for (int i = 0; i < choices_no1; i++)
                {
                    reader1.ReadToFollowing("Data");

                    reader1.MoveToContent();

                    questions[s].values[i] = (questions[s].values[i] + reader1.ReadElementContentAsInt());

                }










                player.redata(ref s, ref questions, ref choices_no1, ref data_format);




            }


            reader.Close();
            reader1.Close();

            player.objApp.ActivePresentation.Save();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            questions[comboBox1.SelectedIndex].back_image = openFileDialog1.FileName;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked == true)
            {
                questions[comboBox1.SelectedIndex].orien = true;

            }
            else
                questions[comboBox1.SelectedIndex].orien = false;

        }

        private void button6_Click(object sender, EventArgs e)
        {
            questions.Capacity=200;
            XmlTextReader reader = new XmlTextReader("c:\\users\\ali\\desktop\\pharma game egypt july.xml");
          //  XmlTextReader reader = new XmlTextReader(folderBrowserDialog1.SelectedPath + "\\" + textBox1.Text + ".xml");
            //XmlTextReader reader1 = new XmlTextReader("c:\\users\\ali\\desktop\\test.xml");
            for (int i = 0; i < 30; i++)
            {
                teams.Add(new Team());
            }

            for (int i = 0; i <= question_no; i++)
            {
                reader.ReadToFollowing("Question");
            }
            if (!reader.EOF)
            {
                questions.Add(new question());
                reader.MoveToAttribute(1);
                questions[question_no].title = reader.Value;

                reader.MoveToAttribute(5);
                questions[question_no].choices_no = reader.Value;
                choices_no1 = Convert.ToInt32(questions[question_no].choices_no);
                reader.MoveToAttribute("CorrectItem");
               questions[question_no].correct_item = Convert.ToInt32(reader.Value);
                //reader.ReadToFollowing("Data"); 
               reader.MoveToAttribute(7);
               fastest = reader.Value;
                int j = 0;
                for (int i = 0; i < choices_no1; i++)
                {
                    
                    reader.ReadToFollowing("Data");
                   reader.MoveToAttribute(1);
                    for (int s = 0; s < reader.Value.Length&&j<30; s++)
                    {
                        if (reader.Value[s] != ',')
                        {
                            teams[j].KeyPad = (teams[j].KeyPad + reader.Value[s]);
                            
                        }
                        
                        else

                        {
                             if ((i + 1) == questions[question_no].correct_item)
                            {
                                teams[j].points=(teams[j].points+10);
                                 
                             }

                            teams[j].bc = Color.FromArgb(random.Next(0, 255), random.Next(0, 255), random.Next(0, 255));
                            j++;
                            
                            
                        }
                            
                            
                        

                    }


                }
                
                   int g=0;
                   string temp1="";
                    string temp2="";
                 for (int s = 0; s <fastest.Length&&g<3; s++)
                    {
                        if (fastest[s] != ',' && fastest[s] != '^')
                        {

                            temp1 = (temp1 + fastest[s]);
                            
                        }


                        else if (fastest[s] == '^')
                        {
                            for (int k = 0; k < 30; k++)
                            {
                                if (teams[k].KeyPad == temp1)
                                {
                                    teams1.Add(new Team());
                                    teams1[g] = teams[k];

                                    s++;
                                    break;
                                }
                            }
                            while (s<fastest.Length&&fastest[s] != ',')
                            {
                                temp2 = (temp2 + fastest[s]);
                                s++;
                            }
                            teams1[g].time = (float.Parse(temp2) / 1000);

                            
                        }
                          if(s<fastest.Length&&fastest[s]==',')
                        {
                            
                            g++;
                            temp1 = "";
                            temp2 = "";
                        }
                  }
                
                //teams.Sort(delegate(Team p1, Team p2) { return p1.KeyPad.CompareTo(p2.KeyPad); });
                 teams1.Sort(delegate(Team p1, Team p2) { return p2.time.CompareTo(p1.time); });
                 teams.OrderBy(t => t.KeyPad).ThenBy(t => t.points);
                 //teams.Sort(delegate(Team p1, Team p2) { return p2.points.CompareTo(p1.points); });
                
                player.add_slide1(ref question_no, ref questions, ref openFileDialog1);
                actHook1 = new UserActivityHook(false, true);
                actHook1.KeyPress += new KeyPressEventHandler(key_pressed1);
                
                player.add_chart1(ref teams,ref teams1);
                teams1.Clear();
               
                player.start_show1();
              
                question_no++;
                                
            }








            reader.Close();



        }
        public void updt1()
        {

            XmlTextReader reader = new XmlTextReader("C:\\Users\\Ali\\Desktop\\pharma game egypt july.xml");
            //XmlTextReader reader1 = new XmlTextReader("c:\\users\\ali\\desktop\\test.xml");


            for (int i = 0; i <= question_no; i++)
            {
                reader.ReadToFollowing("Question");
            }
            if (!reader.EOF)
            {

                questions.Add(new question());
                string temp = "";
                reader.MoveToAttribute(1);
                questions[question_no].title = reader.Value;

                reader.MoveToAttribute(5);
                questions[question_no].choices_no = reader.Value;
                choices_no1 = Convert.ToInt32(questions[question_no].choices_no);
                reader.MoveToAttribute("CorrectItem");
                questions[question_no].correct_item = Convert.ToInt32(reader.Value);
                reader.MoveToAttribute(7);
                fastest = reader.Value;
                //reader.ReadToFollowing("Data"); 
             
                for (int i = 0; i < choices_no1; i++)
                {
                    temp="";
                    reader.ReadToFollowing("Data");
                    reader.MoveToAttribute(1);
                    for (int s = 0; s < reader.Value.Length; s++ )
                    {
                        if (reader.Value[s] != ',')
                        {
                            temp = (temp + reader.Value[s]);

                        }

                        else if ((i + 1) == questions[question_no].correct_item)
                        {
                            int y = 0;

                            for (int h = 0; h < 30; h++)
                            {

                                if (teams[h].KeyPad == temp)
                                {
                                    teams[h].points =( teams[h].points + 10);
                                    y = 1;
                                    temp = "";
                                    break;
                                }
                            }
                            if ((check_empty(ref teams))<= 29 && y==0)
                            { 
                                int z = check_empty(ref teams);
                            teams[z].KeyPad = temp;
                            teams[z].points = teams[z].points + 10;

                            teams[z].bc = Color.FromArgb(random.Next(0, 255), random.Next(0, 255), random.Next(0, 255));
                            temp = "";
                            }

                        }

                        
                         else
                        {
                            int y = 0;

                            for (int h = 0; h < 30; h++)
                            {

                                if (teams[h].KeyPad == temp)
                                {
                                    y = 1;
                                    temp = "";
                                    break;
                                }
                            }
                            if (y == 0)
                            {
                                int z = check_empty(ref teams);
                                teams[z].KeyPad = temp;
                                teams[z].bc = Color.FromArgb(random.Next(0, 255), random.Next(0, 255), random.Next(0, 255));
                                temp = "";
                            }
                        }





                    }


                }

                int g = 0;
                string temp1 = "";
                string temp2 = "";
                teams1.Add(new Team());
                teams1.Add(new Team());
                teams1.Add(new Team());
                for (int s = 0; s < fastest.Length && g < 3; s++)
                {
                    if (fastest[s] != ',' && fastest[s] != '^')
                    {

                        temp1 = (temp1 + fastest[s]);

                    }


                    else if (fastest[s] == '^')
                    {
                        for (int k = 0; k < 30; k++)
                        {
                            if (teams[k].KeyPad == temp1)
                            {

                                teams1[g] = teams[k];

                                s++;
                                while (s < fastest.Length && fastest[s] != ',')
                                {
                                    temp2 = (temp2 + fastest[s]);
                                    s++;
                                }
                                teams1[g].time = (float.Parse(temp2) / 1000);
                                break;
                            }

                        }




                    }
                    if (s < fastest.Length && fastest[s] == ',')
                    {

                        g++;
                        temp1 = "";
                        temp2 = "";
                    }
                }

                // teams.Sort(delegate(Team p1, Team p2) { return p1.points.CompareTo(p2.points); });
                teams.OrderBy(t => t.points).ThenBy(t => t.KeyPad);
                 teams1.Sort(delegate(Team p1, Team p2) { return p2.time.CompareTo(p1.time); });
                player.add_chart1(ref teams, ref teams1);
                teams1.Clear();
                player.start_show1();
                question_no++;


           



            }
















            reader.Close();

        }
         int check_empty(ref List<Team> teams)
        {
            for (int i = 0; i < 30; i++)
            {
                if (teams[i].KeyPad == "")
                    return i;
            }
            return 30;
        }
        
        


    }
}







     
