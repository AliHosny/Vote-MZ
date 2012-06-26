using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
using excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication9
{
    class create_new_pres
    {
        int n = 0;
        int y = 0;
     public   PowerPoint.Application objApp;
        PowerPoint.Presentations objPresSet;
        PowerPoint._Presentation objPres;
        PowerPoint.Slides objSlides;
        PowerPoint._Slide objSlide;
        PowerPoint.TextRange objTextRng;
       
        PowerPoint.Shapes objShapes;
        PowerPoint.Shape objShape;
        PowerPoint.Shape objShape1;
        PowerPoint.Shape objShape2;
        PowerPoint.SlideShowWindows objSSWs;
        PowerPoint.SlideShowTransition objSST;
        PowerPoint.SlideShowSettings objSSS;
        PowerPoint.SlideRange objSldRng;
        Graph.Chart objChart;
        Graph.Chart objChart1;
        Graph.DataSheet datasheet;
        Graph.DataSheet datasheet1;
        public int flag;
        Graph.Points kj;
        PowerPoint._Slide objslide1;
        Graph.Chart objchart2;
        Graph.DataSheet datasheet3;
        Graph.Range range;
        //Create a new and initialize it;s presentation
        public void Init()
        {
            
            objApp = new PowerPoint.Application();
            
            flag = 0;


            objPresSet = objApp.Presentations;
            //objPres = objPresSet.Add(MsoTriState.msoTrue);
          
           objPres= objPresSet.Open("c://sample.pptx");
            objSlides = objPres.Slides;
           
        }
        public void add_title(ref List<question> questions, ref int question_no){
        if (questions[question_no].show_title == true)
            {
                objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
                objTextRng.Text = questions[question_no].title;
                objTextRng.Font.Name = questions[question_no].title_font.Name;
                objTextRng.Font.Color.RGB = ColorTranslator.ToOle(questions[question_no].Title_color);
                objTextRng.Font.Size = questions[question_no].title_font.Size;
                //if (questions[question_no].title_font.Bold == true)
                //    objTextRng.Font.Bold = MsoTriState.msoTrue;
                //else if (questions[question_no].title_font.Bold == false)
                //    objTextRng.Font.Bold = MsoTriState.msoFalse;
               
                //if (questions[question_no].title_font.Underline == true)
                //    objTextRng.Font.Underline = MsoTriState.msoTrue;
                //else if (questions[question_no].title_font.Underline == false)
                //    objTextRng.Font.Underline = MsoTriState.msoFalse;
                //if (questions[question_no].title_font.Italic == true)
                //    objTextRng.Font.Italic = MsoTriState.msoTrue;
                //else if (questions[question_no].title_font.Italic == false)
                //    objTextRng.Font.Italic = MsoTriState.msoFalse;
              




            }
        
        }
        public void add_choices(ref List<question> questions, ref int question_no, ref int choices_no1)
        {
            if (questions[question_no].show_choices == true)
            {
                objShapes = objSlide.Shapes;
                objShape1 = objSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 20, 160, 600, 400);
                objShape1.TextEffect.FontName = questions[question_no].choices_font.Name;
                objShape1.TextEffect.FontSize = questions[question_no].choices_font.Size;
                if (questions[question_no].choices_font.Italic == true)
                    objShape1.TextEffect.FontItalic = MsoTriState.msoTrue;
                else if (questions[question_no].choices_font.Italic == false)
                    objShape1.TextEffect.FontItalic = MsoTriState.msoFalse;
                if (questions[question_no].choices_font.Bold == true)
                    objShape1.TextEffect.FontBold = MsoTriState.msoTrue;
                else if (questions[question_no].choices_font.Bold == false)
                    objShape1.TextEffect.FontBold = MsoTriState.msoFalse;
                objShape1.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(questions[question_no].choices_color); 
               
                objShape1.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                for (int i = 2; i <= choices_no1 + 1; i++)
                {

                    objShape1.TextEffect.Text = objShape1.TextEffect.Text + (i - 1) + "-" + questions[question_no].choices[i - 2] + "\n\n";
                    
                    //if (questions[question_no].choices_font.Italic == true)
                    //    objShape1.TextEffect.FontItalic = MsoTriState.msoTrue;
                    //else if (questions[question_no].choices_font.Italic == false)
                    //    objShape1.TextEffect.FontItalic = MsoTriState.msoFalse;
                    //if (questions[question_no].choices_font.Bold == true)
                    //    objShape1.TextEffect.FontBold = MsoTriState.msoTrue;
                    //else if (questions[question_no].choices_font.Bold == false)
                    //    objShape1.TextEffect.FontBold = MsoTriState.msoFalse;
                    


                }
            }
        }
        public void add_chart(ref List<question> questions, ref int question_no, ref int choices_no1,ref ColorDialog colordialog1, ref int data_format, ref Color[] bar_colors, ref FontDialog fontdialog1)
        {
            if (questions[question_no].show_chart == true)
            {



                objChart = objSlide.Shapes[2].OLEFormat.Object;
                Graph.SeriesCollection ss = objChart.SeriesCollection();
                Graph.Points kj;
                kj = ss.Item(1).Points();
                Graph.Axes kk;
                kk = objChart.Axes();
              
               
               questions[question_no].datasheet = objChart.Application.DataSheet;

                /*
               datasheet.Cells[1, 3] = "ali";
               datasheet.Cells[1, 4] = "ali";
               datasheet.Cells[1, 5] = "ali";
               datasheet.Cells[2, 4] = kk.Count;
              */


                objChart.ApplyDataLabels();
               


                for (int i = 2; i <= choices_no1 + 1; i++)
                {
                    questions[question_no].datasheet.Cells[1, i] = i - 1;

                    // kj.Item(i - 1).Interior.Color = 


                }
                float total = 0;
                for (int i = 2; i <= choices_no1 + 1; i++)
                {
                    total = (total + questions[question_no].values[i - 2]);
                }



                if (data_format == 2)
                {
                    for (int i = 2; i <= choices_no1 + 1; i++)
                    {
                       Graph.DataLabel label = ss.Item(1).DataLabels(i - 1);
                        questions[question_no].datasheet.Cells[2, i] = questions[question_no].values[i - 2];
                        kj.Item(i - 1).Interior.Color = bar_colors[i - 2];
                      

                        
                    }
                    objShape2 = objSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 435, 480, 120, 70);
                    objShape2.TextEffect.Text = "Total" + " " + total.ToString();
                    objShape2.TextEffect.FontName = "Comic Sans MS";
                    objShape2.TextEffect.FontSize = 15;
                    objShape2.TextEffect.FontBold = MsoTriState.msoTrue;
                }
                    
                else if (data_format == 1)
                {
                    for (int i = 2; i <= choices_no1 + 1; i++)
                    {
                        Graph.DataLabel label = ss.Item(1).DataLabels(i - 1);
                        questions[question_no].datasheet.Cells[2, i] = (((questions[question_no].values[i - 2]) / total) * 100) + "%";
                        kj.Item(i - 1).Interior.Color = bar_colors[i - 2];
                        
                    }
                }
                else if (data_format == 3)
                {
                    for (int i = 2; i <= choices_no1 + 1; i++)
                    {
                        questions[question_no].datasheet.Cells[2, i] = questions[question_no].values[i - 2];
                        Graph.DataLabel label = ss.Item(1).DataLabels(i - 1);
                        label.Text = (((questions[question_no].values[i - 2]) / total) * 100).ToString("F") + "%" + "[" + questions[question_no].values[i - 2].ToString() + "]";
                        kj.Item(i - 1).Interior.Color = bar_colors[i - 2];
                        label.Font.Size = 15;
                       
                    }
                    objShape2 = objSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 435, 480, 120, 70);
                    objShape2.TextEffect.Text = "Total" + " " + total.ToString();
                    objShape2.TextEffect.FontName = "Comic Sans MS";
                    objShape2.TextEffect.FontSize = 15;
                    objShape2.TextEffect.FontBold = MsoTriState.msoTrue;
                }




                objChart.Refresh();
                objChart.Application.Update();
         
            }
        }
        public void add_slide( ref int question_no,ref List<question> questions,ref OpenFileDialog open)
        {
            if(questions[question_no].orien==false)
                objSlides.InsertFromFile("c:\\tmphor.ppt", question_no, 1, 1);
            else
          
            objSlides.InsertFromFile("c:\\tmp.ppt", question_no,1,1);
            objSlide = objSlides[question_no + 1];
            objSlide.FollowMasterBackground = MsoTriState.msoFalse;
            objSlide.Background.Fill.BackColor.RGB =ColorTranslator.ToOle( questions[question_no].background);
            if (questions[question_no].back_image != "")
            objSlide.Background.Fill.UserPicture(questions[question_no].back_image);
           
        }

        public void add_slide1(ref int question_no, ref List<question> questions, ref OpenFileDialog open)
        {

            objSlide = objPres.Slides[1];

            objslide1 = objPres.Slides[2];
           //objSlides.InsertFromFile("C:\\Users\\Ali\\Desktop\\Fastest Responders (in seconds).pptx", 0, 1, 1);
           
        
        }
        public void add_chartup(ref List<Team> teams, ref List<Team> teams1)
        {
            for (int i = 2; (i - 2) < 30; i++)
            {
                datasheet.Cells[1, i] = teams[i - 2].KeyPad;

                // kj.Item(i - 1).Interior.Color = 


            }
            for (int i = 2; (i - 2) < 30; i++)
            {
                // Graph.DataLabel label = ss.Item(1).DataLabels(i - 1);
                datasheet.Cells[2, i] = teams[i - 2].points;
                kj.Item(i - 1).Interior.Color = teams[i - 2].bc;



            }
            for (int i = 2; (i - 2) < 3; i++)
            {
                datasheet1.Cells[1, i] = teams1[i - 2].KeyPad;

                // kj.Item(i - 1).Interior.Color = 


            }
            for (int i = 2; (i - 2) < 3; i++)
            {
                // Graph.DataLabel label = ss.Item(1).DataLabels(i - 1);
                datasheet1.Cells[2, i] = teams1[i - 2].time;
                kj.Item(i - 1).Interior.Color = teams1[i - 2].bc;



            }
        
        }

        public void add_chart1(ref List<Team> teams,ref List<Team> teams1 )
        {

            objChart1 = objslide1.Shapes[1].OLEFormat.Object;
            objChart = objSlide.Shapes[1].OLEFormat.Object;
            Graph.SeriesCollection ss = objChart.SeriesCollection();
            
          kj = ss.Item(1).Points();

          
         
         // sc.Item(1).Values = 5;
          
              //  questions[question_no].datasheet = objChart.Application.DataSheet;
            
                /*
               datasheet.Cells[1, 3] = "ali";
               datasheet.Cells[1, 4] = "ali";
               datasheet.Cells[1, 5] = "ali";
               datasheet.Cells[2, 4] = kk.Count;
              */


                //objChart.ApplyDataLabels();

            datasheet= objChart.Application.DataSheet;

           
            
                //float total = 0;
                //for (int i = 2; i <= choices_no1 + 1; i++)
                //{
                //    total = (total + questions[question_no].values[i - 2]);
                //}

            objChart.ApplyDataLabels();



            for (int i = 2; (16 - i) >= 0; i++)
            {
                datasheet.Cells[1, i] = teams[16 - i].KeyPad;

                // kj.Item(i - 1).Interior.Color = 


            }
            for (int i = 2; (16-i) >=0 ; i++)
            {
               // Graph.DataLabel label = ss.Item(1).DataLabels(i - 1);
                datasheet.Cells[2, i] = teams[16-i].points;
                kj.Item(i - 1).Interior.Color = teams[16-i].bc;



            }
            Graph.Axes chart1 = objChart.Axes();
            objchart2 = objSlide.Shapes[2].OLEFormat.Object;
            Graph.Axes chart2 = objchart2.Axes();
            
            ss = objchart2.SeriesCollection();

            kj = ss.Item(1).Points();


            // sc.Item(1).Values = 5;

            //  questions[question_no].datasheet = objChart.Application.DataSheet;

            /*
           datasheet.Cells[1, 3] = "ali";
           datasheet.Cells[1, 4] = "ali";
           datasheet.Cells[1, 5] = "ali";
           datasheet.Cells[2, 4] = kk.Count;
          */


            //objChart.ApplyDataLabels();

            datasheet3 = objchart2.Application.DataSheet;



            //float total = 0;
            //for (int i = 2; i <= choices_no1 + 1; i++)
            //{
            //    total = (total + questions[question_no].values[i - 2]);
            //}

            objchart2.ApplyDataLabels();

            

            for (int i = 2; (31-i) >14; i++)
            {
                datasheet3.Cells[1, i] = teams[31-i].KeyPad;

                // kj.Item(i - 1).Interior.Color = 


            }
            for (int i = 2; (31 - i) >14; i++)
            {
                // Graph.DataLabel label = ss.Item(1).DataLabels(i - 1);
                datasheet3.Cells[2, i] = teams[31 - i].points;
                kj.Item(i - 1).Interior.Color = teams[31 - i].bc;



            }
           
            //second slide
            objChart1 =objslide1.Shapes[1].OLEFormat.Object;
         ss = objChart1.SeriesCollection();
     

          
            // sc.Item(1).Values = 5;

            //  questions[question_no].datasheet = objChart.Application.DataSheet;

            /*
           datasheet.Cells[1, 3] = "ali";
           datasheet.Cells[1, 4] = "ali";
           datasheet.Cells[1, 5] = "ali";
           datasheet.Cells[2, 4] = kk.Count;
          */


            //objChart.ApplyDataLabels();

            datasheet1 = objChart1.Application.DataSheet;



            //float total = 0;
            //for (int i = 2; i <= choices_no1 + 1; i++)
            //{
            //    total = (total + questions[question_no].values[i - 2]);
            //}

            objChart1.ApplyDataLabels();

            

            for (int i = 2; (i - 2) < teams1.Count; i++)
            {
                datasheet1.Cells[1, i] = teams1[i - 2].KeyPad;

                // kj.Item(i - 1).Interior.Color = 


            }
            for (int i = 2; (i - 2) < teams1.Count; i++)
            {
                // Graph.DataLabel label = ss.Item(1).DataLabels(i - 1);
                datasheet1.Cells[2, i] = teams1[i - 2].time;
                kj = ss.Item(1).Points();

                kj.Item(i - 1).Interior.Color = teams1[i - 2].bc;


                
            }

           
           
             




                
            
        }
        
        public void redata(ref int s, ref List<question> questions, ref int choices_no1, ref int data_format)
        {
            float total = 0;
            Graph.SeriesCollection ss = questions[s].datasheet.Application.Chart().SeriesCollection();
            Graph.Points kj;
            kj = ss.Item(1).Points();
            objSlide = objPres.Slides._Index(s + 1);
            objShapes = objSlide.Shapes;
              
            
            for (int i = 2; i <= choices_no1 + 1; i++)
            {
                total = (total + questions[s].values[i - 2]);
            }



            if (data_format == 2)
            {
                for (int i = 2; i <= choices_no1 + 1; i++)
                {
                    
                    questions[s].datasheet.Cells[2, i] = questions[s].values[i - 2];

                   
                }

              questions[s].datasheet.Application.Chart().Refresh();
             
           
            objShape2 = objShapes[4];
            objShape2.TextEffect.Text = "Total" + " " + total.ToString();
            }

            else if (data_format == 1)
            {
                for (int i = 2; i <= choices_no1 + 1; i++)
                {
                  
                    questions[s].datasheet.Cells[2, i] = (((questions[s].values[i - 2]) / total) * 100) + "%";
                  
                   
                }

            }
            else if (data_format == 3)
            {
                for (int i = 2; i <= choices_no1 + 1; i++)
                {
                    questions[s].datasheet.Cells[2, i] = questions[s].values[i - 2];
                    Graph.DataLabel label = ss.Item(1).DataLabels(i - 1);
                   label.Text = (((questions[s].values[i - 2]) / total) * 100).ToString() + "%" + "[" + questions[s].values[i - 2].ToString() + "]";
                   

                  
                }
              
                objShape2 = objShapes[4];
                objShape2.TextEffect.Text = "Total" + " " + total.ToString();
               
            }
        }
           
        public void start_show(ref int question_no)
        {
            objSSWs = objApp.SlideShowWindows;

            objSlide.SlideShowTransition.AdvanceOnClick = MsoTriState.msoTrue;
           
            objApp.PresentationClose+= new Microsoft.Office.Interop.PowerPoint.EApplication_PresentationCloseEventHandler(close_pres);
            // crate an instance with global hooks
            // hang on events
            
            
            //  actHook.KeyPress += new KeyPressEventHandler(key_pressed);
            objSSS = objPres.SlideShowSettings;
            objSSS.StartingSlide = (question_no + 1);
            objSSS.EndingSlide = (question_no + 1);
            objSSS.AdvanceMode = PowerPoint.PpSlideShowAdvanceMode.ppSlideShowManualAdvance;


            objSSS.Run();
            objPres.SlideShowWindow.View.GotoSlide(question_no + 1);
            
            

        }
        public void start_show1()
        {
            objSSWs = objApp.SlideShowWindows;

            objSlide.SlideShowTransition.AdvanceOnClick = MsoTriState.msoTrue;

            objApp.PresentationClose += new Microsoft.Office.Interop.PowerPoint.EApplication_PresentationCloseEventHandler(close_pres);
            // crate an instance with global hooks
            // hang on events


            //  actHook.KeyPress += new KeyPressEventHandler(key_pressed);
            objSSS = objPres.SlideShowSettings;
            objSSS.StartingSlide = (1);
            objSSS.EndingSlide = (1);
            objSSS.AdvanceMode = PowerPoint.PpSlideShowAdvanceMode.ppSlideShowManualAdvance;


            objSSS.Run();
            objPres.SlideShowWindow.View.GotoSlide(1);



        }
        private void close_pres(PowerPoint.Presentation pres)
        {
           
            flag = 1;
            
           
            
        }
    }
}
