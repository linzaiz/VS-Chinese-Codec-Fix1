

/*   VS中文乱码修复程序
 * VS2010的代码粘贴到Word里面的汉字乱码修正问题
2014年11月17日 17:31:20 冰殇曙 阅读数 728    https://blog.csdn.net/u012256422/article/details/41212671
直接复制VS2010的代码到Word里面去时，汉字会出现乱码，
虽然可以采用记事本打开然后复制的方法，但是这样就失去了高亮色，不是我们想要的，
下面的小程序就是解决这个问题的。

使用时，先正常复制代码到剪贴板，然后点击“乱码修正”按钮，最后直接粘贴到Word里面就不会有乱码了。
*/
using System;

using System.Collections.Generic;

using System.ComponentModel;

using System.Data;

using System.Drawing;

using System.Text;
using System.Windows.Forms;

namespace VS中文乱码修复1
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.button乱码修正 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button乱码修正
            // 
            this.button乱码修正.Location = new System.Drawing.Point(115, 46);
            this.button乱码修正.Name = "button乱码修正";
            this.button乱码修正.Size = new System.Drawing.Size(75, 23);
            this.button乱码修正.TabIndex = 0;
            this.button乱码修正.Text = "button乱码修正";
            this.button乱码修正.UseVisualStyleBackColor = true;
            this.button乱码修正.Click += new System.EventHandler(this.button乱码修正_Click_1);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.button乱码修正);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        /*
        private void button乱码修正_Click(object sender, EventArgs e)

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
        */
        private Button button乱码修正;
        
    }
}

