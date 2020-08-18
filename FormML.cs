using System;
using System.Collections.Generic;
using System.ComponentModel;
using sysData = System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.AutoCAD.ApplicationServices;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.IO;
//using ExcelDataReader;
//using np = NPOI.XSSF.UserModel;
using Autodesk.AutoCAD.DatabaseServices;

namespace ML
{
    public partial class FormML : Form
    {
        public FormML()
        {
            InitializeComponent();
        }

        private void FormML_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";

            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "文本文件(*.txt)|*.txt";
            file.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            file.Multiselect = false;
            DialogResult result = file.ShowDialog();
            if (result == DialogResult.Cancel)
                return;
            else if(result == DialogResult.OK || result == DialogResult.Yes)
            {
                textBox1.Text = file.FileName;
            }

        }
       
        private void button2_Click(object sender, EventArgs e)
        {
            Document doc = acadApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            
            if (textBox1.Text == "" | textBox2.Text == "" | textBox3.Text == "" | textBox4.Text == "" | textBox5.Text == "")
            {
                MessageBox.Show("请将所有选项填写齐全！！！");
            }
            else
            {
                
                string mlPath = textBox1.Text;
                string proNo = textBox2.Text;
                string proName = textBox3.Text;
                string proJD = comboBox1.Text;
                string sK = textBox4.Text;
                string zB = textBox5.Text;

                this.Hide();                

                //提取.txt文件中的数据
                string contents = "";
                StreamReader streamReader = new StreamReader(mlPath);
                //contents = streamReader.ReadToEnd();
                do
                {
                    contents = contents + streamReader.ReadLine() + "\n";
                } while (streamReader.Peek() != -1);
                streamReader.Close();

                string[] nameAndNo = contents.Split(new char[] { '\n', '\t' });
                int dwgnums;//计算图纸页数
                if ((nameAndNo.Length / 2) % 21 != 0)
                    dwgnums = ((nameAndNo.Length / 2) / 21) + 1;
                else
                    dwgnums = (nameAndNo.Length / 2) / 21;
                
                DateTime dt = DateTime.Now;
                
                ObjectId spaceId = db.CurrentSpaceId;//获取当前空间（模型空间或图纸空间）

                //图框插入起始点
                PromptPointOptions ppo1 = new PromptPointOptions("\n给定目录图框起始点：");
                PromptPointResult ppr1 = ed.GetPoint(ppo1);
                if (ppr1.Status == PromptStatus.OK)
                {
                    Point3d spt = ppr1.Value;
                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        string dwgFileName = ComLibClass.GetCurrentPath() + "\\图块文件.dwg";
                        string blkname = "目录图框";
                        int x = 2;
                        for (int i = 1; i <= dwgnums; i++)
                        {
                            db.ImportBlocksFromDWG(dwgFileName, blkname);
                            //表示属性的字典对象
                            Dictionary<string, string> atts = new Dictionary<string, string>();
                            atts.Add("PAGE", i.ToString());
                            atts.Add("PAGES", dwgnums.ToString());
                            atts.Add("GCH", proNo);
                            atts.Add("JD", proJD);
                            atts.Add("PROJECTNMAE", proName);
                            atts.Add("DIRECTOR", sK);
                            atts.Add("CTO", zB);
                            atts.Add("N", dt.ToString("yy"));
                            atts.Add("Y", dt.ToString("MM"));
                            atts.Add("R", dt.ToString("dd"));                            
                            spaceId.InsertBlockReference("0", blkname, spt, new Scale3d(1), 0, atts);
                            Point3d insertPoint = new Point3d(spt.X + 3300, spt.Y + 21540, 0);
                            for (int n = 0; n < 21; n++)
                            {
                                if (x >= nameAndNo.Length-2) break;
                                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                                {
                                    // Open the Block table for read以读打开Block表
                                    BlockTable acBlkTbl;
                                    acBlkTbl = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                                    // Open the Block table record Modelspace for write以写打开块表记录模型空间
                                    BlockTableRecord acBlkTblRec;
                                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                    // Create a single-line text object创建一个单行文字对象
                                    DBText acText1 = new DBText();
                                    acText1.Position = insertPoint;
                                    acText1.Height = 300;
                                    acText1.TextString = nameAndNo[x];
                                    x++;
                                    // Change the oblique angle of the textobject to 45 degrees(0.707 in radians)修改文字对象的倾角为45度（弧度值为0.707）
                                    //acText.Oblique = 0.707;
                                    acBlkTblRec.AppendEntity(acText1);
                                    acTrans.AddNewlyCreatedDBObject(acText1, true);

                                    DBText acText2 = new DBText();
                                    acText2.Position = new Point3d(insertPoint.X + 3400, insertPoint.Y, insertPoint.Z);
                                    acText2.Height = 300;
                                    acText2.TextString = nameAndNo[x];
                                    x++;
                                    acBlkTblRec.AppendEntity(acText2);
                                    acTrans.AddNewlyCreatedDBObject(acText2, true);

                                    DBText acText3 = new DBText();
                                    acText3.Position = new Point3d(insertPoint.X + 10900, insertPoint.Y, insertPoint.Z);
                                    acText3.Height = 300;
                                    acText3.TextString = dt.ToString("yyyy-MM-dd");
                                    acBlkTblRec.AppendEntity(acText3);
                                    acTrans.AddNewlyCreatedDBObject(acText3, true);
                                    // Save the changes and dispose of thetransaction保存修改，关闭事务
                                    acTrans.Commit();
                                }
                                
                                insertPoint = new Point3d(insertPoint.X, insertPoint.Y - 780, insertPoint.Z);
                            }
                            spt = new Point3d(spt.X + 21000, spt.Y, spt.Z);
                        }
                        trans.Commit();
                    }
                    this.Close();
                }
                else return;
            }            
        }
    }    
}
