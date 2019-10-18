using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using pptx = Microsoft.Office.Interop.PowerPoint;
using docx = Microsoft.Office.Interop.Word;

namespace Export_img_txtbox_memo_in_pptx_to_docx_WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //預設按鈕
            this.AcceptButton = button1;//https://docs.microsoft.com/zh-tw/dotnet/framework/winforms/controls/how-to-designate-a-windows-forms-button-as-the-accept-button
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //const string s = @"C:\Users\oscar\OneDrive\公用\mis2000lab-ASP\公用\2016_台中科大\WebService_REST精簡版2.pptx";
            string s = textBox1.Text; //@"file:///C:\Users\oscar\Dropbox\「自傳」簡報大綱.pptx";
            if (!System.IO.File.Exists(s.Replace("file:///", "").Replace("%20"," ")))
            {
                MessageBox.Show("檔案不存在，請重新操作！", "！！檔案全名（含路徑）有誤！！", MessageBoxButtons.OK, MessageBoxIcon.Error); return;
            };
            docx.Application w = new docx.Application();
            docx.Document d = w.Documents.Add();
            d.ActiveWindow.Visible = true;
            docx.Selection slct = d.ActiveWindow.Selection;
            //if (d.Path!="")
            //{
            //    if (MessageBox.Show("不是新檔，確定繼續？","11", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk) == DialogResult.Cancel) return;
            //}
            pptx.Application papp= new pptx.Application();
            pptx.Presentations p = papp.Presentations;
            pptx.Presentation activePresentation = p.Open(s);
            string txt = "";// shapeType = "";
            foreach (pptx.Slide sl in activePresentation.Slides)
            {
                foreach (pptx.Shape sp in sl.Shapes)
                {
                    //shapeType = sp.Name;
                    //if (shapeType.IndexOf("Text") == -1 &&
                    //    shapeType.IndexOf("Title") == -1)//not textframe
                    //                                     //TextBox 、"Text Placeholder"、"Rectangle"
                    if (sp.HasTextFrame == 0)//not textframe//https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff746607(v=office.14) 
                    {//判斷含不含文字方塊是用Shape的HasTexFrame這個屬性來判斷！
                        Clipboard.Clear();
                        sp.Copy();
                        try
                        {
                            slct.Paste();

                        }
                        catch (Exception)
                        {
                            sp.Copy();
                            //throw;
                        }
                        slct.Collapse(docx.WdCollapseDirection.wdCollapseEnd);
                    }
                    else
                    {
                        txt = sp.TextFrame.TextRange.Text;
                        slct.TypeText(txt);
                        slct.InsertAfter("\n\r");
                        slct.Collapse(docx.WdCollapseDirection.wdCollapseEnd);
                    }

                }
                if (sl.HasNotesPage == Microsoft.Office.Core.MsoTriState.msoTrue)//if (sl.HasNotesPage==-1)
                {
                    if (sl.NotesPage.Shapes.Count > 0)
                        foreach (pptx.Shape item in sl.NotesPage.Shapes)
                        {
                            if (item.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                if (item.Name.IndexOf("Slide Number")!=-1)
                                    slct.TypeText("Slide Number: " + item.TextFrame.TextRange.Text);
                                else
                                slct.TypeText(item.TextFrame.TextRange.Text);
                                slct.InsertAfter("\n\r");
                                slct.Collapse(docx.WdCollapseDirection.wdCollapseEnd);
                            }
                        }
                }
            }
            activePresentation.Close();
            papp.Quit(); activePresentation = null; p = null; papp = null;            
            MessageBox.Show("簡報內的圖文已順利匯出到Word文件了!","",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            try
            {
                textBox1.Text= Clipboard.GetText();
            }
            catch (Exception)
            {
                return;
                //throw;
            }
        }
    }
}
