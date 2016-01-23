using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;
using iTextSharp.text.pdf;
using ICSharpCode.SharpZipLib.Zip;

namespace WindowsFormsApplication1 {
    public partial class Form1 : Form {

        Microsoft.Office.Interop.Word.Application docApp;
        Microsoft.Office.Interop.Word.Documents docDocs;
        Microsoft.Office.Interop.Word.Document docDoc;

        private void checkWord() {
            if (Type.GetTypeFromProgID("Word.Application") != null) {
                avaWord = true;
                Pic_Word.Visible = true;
            }
        }

        private void addWord(String filePath) {
            if (avaWord) {
                // DataGridへWord追加
                dataGridView1.Rows.Add();
                int idx = dataGridView1.Rows.Count - 1;
                dataGridView1.Rows[idx].Cells[1].Value = Path.GetFileName(filePath);
                dataGridView1.Rows[idx].Cells[2].Value = "Word";
                dataGridView1.Rows[idx].Cells[3].Value = filePath;
            }
        }

        private void allocWord() {
            if (avaWord) {
                docApp = new Microsoft.Office.Interop.Word.Application();
                docDocs = docApp.Documents;
            }
        }

        private void freeWord() {
            if (avaWord) {
                if (docDocs != null)
                    docDocs = null;
                ((Microsoft.Office.Interop.Word._Application)docApp).Quit();
                ;
            }
        }

        private Boolean lockWord(String r_password, String w_password) {
            //////////////////////////////////////////////////////
            // Word施錠
            //////////////////////////////////////////////////////
            if (avaWord) {
                // 読み込みパスワードが設定されているかチェック
                try {
                    docDoc = docDocs.Open(dataGridView1.Rows[0].Cells[3].Value.ToString(),
                        Type.Missing, true, Type.Missing, " ", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false);
                } catch {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.word1;
                    //"This document has been read password-protected.";
                    docDoc = null;
                    return false;
                }
                ((Microsoft.Office.Interop.Word._Document)docDoc).Close();
                // 書き込みパスワードが設定されているかチェック
                try {
                    docDoc = docDocs.Open(dataGridView1.Rows[0].Cells[3].Value.ToString(),
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, " ", Type.Missing, Type.Missing, Type.Missing, false);
                } catch {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.word2;
                    // "This document has been write password-protected.";
                    docDoc = null;
                    return false;
                }
                if (textBox1.Text.Length > 0)
                    docDoc.Password = r_password;
                if (textBox2.Text.Length > 0)
                    docDoc.WritePassword = w_password;
                try {
                    docDoc.Save();
                } catch (Exception eX) {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                    // "Saving failed." + eX.Message;
                    return false;
                }
                docDoc = null;
            }
            return true;
        }

        private Boolean unlockWord(String r_password, String w_password) {
            Boolean NOPASS = true;

            /////////////////////////////////////////////////////////
            // Word解錠
            /////////////////////////////////////////////////////////
            if (avaWord) {
                // 一時ファイル取得
                String tmpFilePath = Path.GetTempFileName();
                try {
                    docDoc = docDocs.Open(dataGridView1.Rows[0].Cells[3].Value.ToString(),
                        Type.Missing, Type.Missing, Type.Missing, " ", Type.Missing, Type.Missing, " ", Type.Missing, Type.Missing, Type.Missing, false);
                } catch {
                    NOPASS = false;
                }
                if (NOPASS) {
                    // パスワードがかかっていない
                    label2.Text = WindowsFormsApplication1.Properties.Resources.word3;
                    // "This document is not applied password.";
                    docDoc = null;
                    return false;
                } else {
                    // パスワード無しの時は、スペース1つ
                    String rp = (r_password.Length == 0) ? " " : r_password;
                    String wp = (w_password.Length == 0) ? " " : w_password;
                    try {
                        docDoc = docDocs.Open(dataGridView1.Rows[0].Cells[3].Value.ToString(),
                            Type.Missing, Type.Missing, Type.Missing, rp, Type.Missing, Type.Missing, wp, Type.Missing, Type.Missing, Type.Missing, false);
                    } catch {
                        label2.Text = WindowsFormsApplication1.Properties.Resources.message2;
                        //"Password is incorrect.";
                        docDoc = null;
                        return false;
                    }

                    docDoc.Password = null;
                    docDoc.WritePassword = null;
                    try {
                        docDoc.SaveAs(tmpFilePath,
                            Type.Missing, Type.Missing, "", Type.Missing, "");
                        ((Microsoft.Office.Interop.Word._Document)docDoc).Close();
                        System.IO.File.Copy(tmpFilePath,
                            dataGridView1.Rows[0].Cells[3].Value.ToString(), true);
                        System.IO.File.Delete(tmpFilePath);
                    } catch (Exception eX) {
                        label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                        // "Saving failed." + eX.Message;
                        return false;
                    }
                }
                docDoc = null;
            }
            return true; 
        }
    }
}
