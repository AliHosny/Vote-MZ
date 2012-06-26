using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
    class Team
    {
        Random random = new Random();
        public int points;
        public string KeyPad;
        public Color bc;
        public float time;
        public Team()
        {
           
            points = 0;
            KeyPad = "";
            bc = Color.FromArgb(random.Next(0, 255), random.Next(0, 255), random.Next(0, 255));
            

        }
    }
}
