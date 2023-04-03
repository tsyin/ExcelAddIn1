using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Core;
using System.Collections;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.AccessControl;
using System.Linq;
using System.Collections.Generic;
using Config;
using System;
using Config;

namespace ExcelAddIn1
{

    public partial class Ribbon1
    {
        //一个动态数组存储地址表格地址
        ArrayList news = new ArrayList();
        struct vecs2
        {
            string cell1_value;
            string cell2_value;
            public void set(string cell1_value, string cell2_value) { this.cell1_value = cell1_value; this.cell2_value = cell2_value;}
        };
        private ArrayList GetRang(string s1,string s2,Excel.Workbooks wb)
        {
            ArrayList air_list = new ArrayList();
            foreach (Excel.Workbook item in wb) 
            {
                Excel.Worksheet ws = item.Worksheets[0];
                int nrows = ws.UsedRange.Rows.Count;
                int ncols = ws.UsedRange.Columns.Count;
                int s1_num = 0,s2_num = 0;
                for (int i = 0; i < ncols; i++)
                {
                    string temp = ws.UsedRange.Rows[0].Cells[i].Value;
                    if(s1 == temp)
                        s1_num = i;
                    else if (s2 == temp)
                        s2_num = i;
                }
                for (int i = 0;i<nrows; i++)
                {
                    vecs2 temp = new vecs2();
                    temp.set(ws.Cells[i, s1_num], ws.Cells[i, s2_num]);
                    air_list.Add(temp);
                }
            }
            return air_list;
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            RibbonCheckBox[] chebox = { checkBox1, checkBox2, checkBox3, checkBox4 };
            string[] checkname = { "东莞", "惠州", "深圳", "汕尾" };
            for (int i = 0;i<chebox.Length;i++)
            {
                chebox[i].Checked = Config.ClsThisAddinConfig.GetInstance().readEle<bool>(checkname[i], false);
                //Debug.WriteLine(chebox[i].Label, chebox[i].Checked.ToString());
            }
        }

        private ArrayList getCheckedBox() 
        {
            RibbonCheckBox[] boxs = { checkBox1, checkBox2, checkBox3, checkBox4 };
            ArrayList list = new ArrayList();
            int count = 0;
            foreach (RibbonCheckBox box in boxs) 
            {
                if (box.Checked)
                {
                    count++;
                    list.Add(box);
                }
            }
            return list;
        }

        //打开表格
        private List<Workbook> OpenFileDialogFile(string title) 
        {
            Application app = Globals.ThisAddIn.Application;
            FileDialog filedialog = app.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
            filedialog.AllowMultiSelect = true;
            filedialog.Filters.Add("03版Excel文件", "*.xls", 1);
            filedialog.Filters.Add("07版Excel文件", "*.xlsx", 2);
            filedialog.Filters.Add("带宏的Excel文件", "*.xlsm", 3);
            filedialog.Title = title;
            List<string> files = new List<string>();
            List<Workbook> wbs = new List<Workbook>();
            if (filedialog.Show() == -1)
            {
                FileDialogSelectedItems fdsi = filedialog.SelectedItems;
                for (int i = 0; i < fdsi.Count; i++)
                {
                    string fdstr = fdsi.Item(i+1);
                    files.Add(fdstr);
                    Debug.WriteLine(fdstr);
                    Workbook temp = Globals.ThisAddIn.Application.Workbooks.Open(fdstr);
                    wbs.Add(temp);
                }
            }
            return wbs;
        }

        //自动生成准备前
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbooks wb = Globals.ThisAddIn.Application.Workbooks;
            Application app = Globals.ThisAddIn.Application;
            int count = wb.Count;
            //获得地市
            ArrayList boxs = getCheckedBox();
            string[] dishi = new string[boxs.Count];
            for (int i = 0; i < dishi.Length; i++)
            {
                RibbonCheckBox tempbox = (RibbonCheckBox)boxs[i];
                dishi[i] = tempbox.Label;
            }
            List<Workbook> wbs = OpenFileDialogFile("打开你需要的准备表");
        }

        //读取当前选中的地市
        private List<string> GetSelectCity() 
        {
            ArrayList chBoxs = this.getCheckedBox();
            List<string> citys = new List<string>();
            foreach (RibbonCheckBox item in chBoxs)
            {
                if (item.Checked)
                    citys.Add(item.Label); 
            }
            return citys;
        }
        
        //生成操作前准备表
        private void CreatePrepareWB()
        {
            //读取当前选中的地市
            List<string> citys = GetSelectCity();
        }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            Config.ClsThisAddinConfig.GetInstance().writeEle("东莞", checkBox1.Checked);
        }

        private void checkBox2_Click(object sender, RibbonControlEventArgs e)
        {
            Config.ClsThisAddinConfig.GetInstance().writeEle("惠州", checkBox2.Checked);
        }

        private void checkBox3_Click(object sender, RibbonControlEventArgs e)
        {
            Config.ClsThisAddinConfig.GetInstance().writeEle("深圳", checkBox3.Checked);
        }

        private void checkBox4_Click(object sender, RibbonControlEventArgs e)
        {
            Config.ClsThisAddinConfig.GetInstance().writeEle("汕尾", checkBox4.Checked);
        }
    }
}
