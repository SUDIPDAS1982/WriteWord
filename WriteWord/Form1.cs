using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace WriteWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

		private void button1_Click(object sender, EventArgs e)
		{
			Object pInfoSheetName = @"D:\SUDIP DAS\Doc1.docx";
			Object pMissing = Missing.Value;
			Word.Application pWordApp = new Word.Application();
			Word.Document pWordDoc = null;
			pWordDoc = pWordApp.Documents.Add(ref pInfoSheetName, ref pMissing, ref pMissing, ref pMissing);			
			pWordDoc.Bookmarks["bmrkName"].Range.Text = "My Name";
			pWordApp.Visible = true;
		}

    }
}
