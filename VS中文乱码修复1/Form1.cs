using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace VS中文乱码修复1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button乱码修正_Click_1(object sender, EventArgs e)
        {

            {

                //取得剪贴板内容

                IDataObject dataObject = Clipboard.GetDataObject();

                if (dataObject.GetDataPresent(DataFormats.Rtf))

                {

                    //取出RTF格式

                    string rtf = dataObject.GetData(DataFormats.Rtf) as string;

                    //以Regex.Replace去除多余字元(註: 不管是否有問題，一律強制處理)

                    string fixedRtf =

                      System.Text.RegularExpressions.Regex.Replace(rtf, @"\\uinput2(?<uc>\\u-?\d*)\s..",

                    (m) =>

                    {

                        return m.Groups["uc"].Value + "?";

                    });

                    //另建新DataObject物件

                    DataObject newDataObject = new DataObject();

                    //RTF格式用修正後的字串，其餘依原值

                    foreach (String t in dataObject.GetFormats())

                        newDataObject.SetData(t,

                        t == "Rich Text Format" ? fixedRtf :

                        dataObject.GetData(t));

                    //将修正内容写入剪贴板

                    Clipboard.SetDataObject(newDataObject, true);

                    MessageBox.Show("中文乱码修正成功!\n现在您可以直接到Word里按Ctrl+V粘贴了！", "成功");

                }

                else

                    MessageBox.Show("您粘贴的不是代码！", "错误");

                }
        }
   }
}
