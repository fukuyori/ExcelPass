using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace WindowsFormsApplication1 {
    public partial class Form1 : Form {

        Microsoft.Office.Interop.Excel.Application xlApp;
        Microsoft.Office.Interop.Excel.Workbooks xlBooks;
        Microsoft.Office.Interop.Excel.Workbook xlBook;

        private void checkExcel() {
            if (Type.GetTypeFromProgID("Excel.Application") != null) {
                avaExcel = true;
                Pic_Excel.Visible = true;
            }
        }

        private void addExcel(String filePath) {
            if (avaExcel) {
                // DataGridへExcel表追加
                dataGridView1.Rows.Add();
                int idx = dataGridView1.Rows.Count - 1;
                dataGridView1.Rows[idx].Cells[1].Value = Path.GetFileName(filePath);
                dataGridView1.Rows[idx].Cells[2].Value = "Excel";
                dataGridView1.Rows[idx].Cells[3].Value = filePath;
            }
        }

        private void allocExcel() {
            if (avaExcel) {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlBooks = xlApp.Workbooks;
            }
        }

        private void freeExcel() {
            if (avaExcel) {
                if (xlBook != null)
                    xlBook = null;
                xlApp.Quit();
            }
        }

        private Boolean lockExcel(String r_password, String w_password) {
            ////////////////////////////////////////////////////////
            // Excel施錠
            ////////////////////////////////////////////////////////
            string orgFileName = dataGridView1.Rows[0].Cells[3].Value.ToString();
            string tmpFileName = Path.GetTempFileName() + Path.GetExtension(orgFileName);

            if (avaExcel) {
                // 読み込みパスワードが設定されているかチェック
                try {
                    xlBook = xlBooks.Open(orgFileName,
                        Type.Missing, true, Type.Missing, "");
                } catch {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.excel1;
                    //"This workbook has been read password-protected.";
                    // xlBooks.Close();
                    return false;
                }

                // 書き込みパスワードが設定されているかチェック
                try {
                    xlBook = xlBooks.Open(orgFileName,
                        Type.Missing, Type.Missing, Type.Missing, "", "");
                } catch {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.excel2;
                    //"This workbook has been write password-protected.";
                    xlBooks.Close();
                    return false;
                }
                // 共有ブックの保護
                if (xlBook.MultiUserEditing) {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.excel4;
                    // "This workbook's permission is set to shared with multi users.";
                    xlBook.Close(false);
                    xlBooks.Close();
                    return false;
                }


                if (textBox1.Text.Length > 0 & textBox2.Text.Length > 0) {
                    // Read Pass and Write Pass
                    try {
                        xlBook.SaveAs(tmpFileName, xlBook.FileFormat, r_password, w_password, xlBook.ReadOnlyRecommended);
                    } catch (Exception eX) {
                        label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                        //"Saving failed." + eX.Message;
                        return false;
                    }
                } else if (textBox1.Text.Length > 0) {
                    // Read Pass
                    try {
                        xlBook.SaveAs(tmpFileName, xlBook.FileFormat, r_password, Type.Missing, xlBook.ReadOnlyRecommended);
                    } catch (Exception eX) {
                        label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                        //"Saving failed." + eX.Message;
                        return false;
                    }
                } else if (textBox2.Text.Length > 0) {
                    // Write Pass
                    try {
                        xlBook.SaveAs(tmpFileName, xlBook.FileFormat, Type.Missing, w_password, xlBook.ReadOnlyRecommended);
                    } catch (Exception eX) {
                        label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                        //"Saving failed." + eX.Message;
                        return false;
                    }
                }

                xlBooks.Close();

                try {
                    System.IO.File.Copy(tmpFileName,orgFileName, true);
                    System.IO.File.Delete(tmpFileName);
                } catch (Exception eX) {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                    // "Saving failed." + eX.Message;
                    return false;
                }
            }
            return true;
        }

        private Boolean unlockExcel(String r_password, String w_password) {
            Boolean NOPASS = true;

            /////////////////////////////////////////////////////////////
            // Excel解錠
            /////////////////////////////////////////////////////////////
            if (avaExcel) {
                try {
                    xlBook = xlBooks.Open(dataGridView1.Rows[0].Cells[3].Value.ToString(),
                        Type.Missing, Type.Missing, Type.Missing, "", "");
                } catch {
                    NOPASS = false;
                }
                if (NOPASS) {
                    // パスワードがかかっていない
                    label2.Text = WindowsFormsApplication1.Properties.Resources.excel3;
                    // "This workbook is not applied password.";
                    xlBooks.Close();
                    return false;
                } else {
                    try {
                        xlBook = xlBooks.Open(dataGridView1.Rows[0].Cells[3].Value.ToString(),
                            Type.Missing, Type.Missing, Type.Missing, r_password, w_password);
                    } catch {
                        label2.Text = WindowsFormsApplication1.Properties.Resources.message2;
                        // "Password is incorrect.";
                        //  xlBooks.Close();
                        return false;
                    }
                    // 共有ブックの保護
                    if (xlBook.MultiUserEditing) {
                        label2.Text = WindowsFormsApplication1.Properties.Resources.excel4;
                        // "This workbook's permission is set to shared with multi users.";
                        xlBook.Close(false);
                        xlBooks.Close();
                        return false;
                    }
                    xlBook.Password = "";
                    xlBook.WritePassword = "";
                    xlBook.CheckCompatibility = false;
                    try {
                        xlBook.Save();
                    } catch (Exception eX) {
                        label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                        // "Saving failed." + eX.Message;
                        return false;
                    }
                    xlBooks.Close();
                }
            }
            return true;
        }
    }
}
