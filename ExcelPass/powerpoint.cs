using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;
using iTextSharp.text.pdf;
using ICSharpCode.SharpZipLib.Zip;

namespace WindowsFormsApplication1 {
    public partial class Form1 : Form {

        Microsoft.Office.Interop.PowerPoint.Application pptApp;
        Microsoft.Office.Interop.PowerPoint.Presentations pptPres;
        Microsoft.Office.Interop.PowerPoint.Presentation pptPre;

        private void checkPowerPoint() {
            if (Type.GetTypeFromProgID("PowerPoint.Application") != null) {
                avaPoawrPoint = true;
                Pic_PowerPoint.Visible = true;
            }
        }

        private void addPowerPoint(String filePath) {
            if (avaPoawrPoint) {
                // DataGridへPowerPoint追加
                dataGridView1.Rows.Add();
                int idx = dataGridView1.Rows.Count - 1;
                dataGridView1.Rows[idx].Cells[1].Value = Path.GetFileName(filePath);
                dataGridView1.Rows[idx].Cells[2].Value = "PowerPoint";
                dataGridView1.Rows[idx].Cells[3].Value = filePath;
            }
        }

        private void allocPowerPoint() {
            if (avaPoawrPoint) {
                pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
                pptPres = pptApp.Presentations;
            }
        }

        private void freePowerPoint() {
            if (avaPoawrPoint) {
                if (pptPres != null)
                    pptPres = null;
                pptApp.Quit();
            }
        }

        private Boolean lockPowerPoint(String r_password, String w_password) {
            ////////////////////////////////////////////////////
            // PowerPoint施錠
            ////////////////////////////////////////////////////
            if (avaPoawrPoint) {
                // パスワードが設定されているかチェック
                try {
                    pptPre = pptPres.Open(dataGridView1.Rows[0].Cells[3].Value.ToString() + ":: :: ",
                        Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoFalse);
                } catch {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.powerpoint1;
                    // "This presentation has been password-protected.";
                    return false;
                }

                if (textBox1.Text.Length > 0)
                    pptPre.Password = r_password;
                if (textBox2.Text.Length > 0)
                    pptPre.WritePassword = w_password;
                try {
                    pptPre.Save();
                } catch (Exception eX) {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                    // "Saving failed." + eX.Message;
                    return false;
                }
                pptPre.Close();
            }
            return true;
        }

        private Boolean unlockPowerPoint(String r_password, String w_password) {
            Boolean NOPASS = true;

            /////////////////////////////////////////////////////
            // PwerPoint解錠
            /////////////////////////////////////////////////////
            if (avaPoawrPoint) {
                try {
                    pptPre = pptPres.Open(dataGridView1.Rows[0].Cells[3].Value.ToString() + ":: :: ",
                        MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                } catch {
                    NOPASS = false;
                }

                if (NOPASS) {
                    // パスワードがかかっていない
                    label2.Text = WindowsFormsApplication1.Properties.Resources.powerpoint2;
                    // "This presentation is not applied password.";
                    return false;
                } else {
                    String rp = (r_password.Length == 0) ? ":: " : "::" + r_password;
                    String wp = (w_password.Length == 0) ? ":: " : "::" + w_password;

                    try {
                        pptPre = pptPres.Open(dataGridView1.Rows[0].Cells[3].Value.ToString() + rp + wp,
                            Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoFalse);
                    } catch {
                        label2.Text = WindowsFormsApplication1.Properties.Resources.message2;
                        //"Password is incorrect.";
                        return false;
                    }

                    pptPre.Password = null;
                    pptPre.WritePassword = null;
                    try {
                        pptPre.Save();
                    } catch (Exception eX) {
                        label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                        // "Saving failed." + eX.Message;
                        return false;
                    }
                    pptPre.Close();
                }
            }
            return true;
        }
    }
}
