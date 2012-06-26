using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System;
using System.Collections.Generic;
using System.Linq;
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
    class question
    {
        public
                  List<String> choices = new List<String>();
        public String title;
        
        public string choices_no;
        public List<float> values = new List<float>();
        public int correct_item = 0;
        public Color Title_color;
        public Font title_font;
        public Color choices_color;
        public Font choices_font;
        public bool show_choices;
        public bool show_title;
        public bool show_chart;
        public  Color background;
        public Graph.DataSheet datasheet;
        public string back_image="";
        public bool orien;
        public string[] right_teams= new string[30];
        public string right;
        public void parse()
        {
            for ( int i=0,j = 0; i < right.Length;i++ )
            {
                if (right[i] != ',')
                {
                    right_teams[j] = (right_teams[0] + right[i]);
                }
                else
                    j++;

            }
        }
        



    }
}
