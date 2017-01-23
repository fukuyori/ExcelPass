using System;
using System.IO;
using System.Windows.Forms;
using iTextSharp.text.pdf;

namespace WindowsFormsApplication1 {
    public partial class Form1 : Form {

        iTextSharp.text.pdf.PdfReader pdfReader;
        iTextSharp.text.pdf.PdfCopy pdfCopy;
        iTextSharp.text.Document pdfDoc;
        FileStream os;

        private void checkPDF() {
            avaPDF = true;
            Pic_Pdf.Visible = true;
        }

        private void addPDF(String filePath) {
            if (avaPDF) {
                // DataGridへPDF追加
                dataGridView1.Rows.Add();
                int idx = dataGridView1.Rows.Count - 1;
                dataGridView1.Rows[idx].Cells[1].Value = Path.GetFileName(filePath);
                dataGridView1.Rows[idx].Cells[2].Value = "PDF";
                dataGridView1.Rows[idx].Cells[3].Value = filePath;
            }
        }

        private void allocPDF() {
        }

        private void freePDF() {
            if (avaPDF) {
                try {
                    pdfReader.Dispose();
                    pdfDoc.Dispose();
                    pdfCopy.Dispose();
                } catch {
                    pdfReader = null;
                    pdfDoc = null;
                    pdfCopy = null;
                }
            }
        }

        private Boolean lockPDF(String r_password, String w_password) {
            string rp = null, wp = null;

            ////////////////////////////////////////////////////
            // PDF施錠
            ////////////////////////////////////////////////////
            if (avaPDF) {
                // 一時ファイル取得
                String tmpFilePath = Path.GetTempFileName();

                // パスワードなしで読み込み可能かチェック
                try {
                    pdfReader = new PdfReader(dataGridView1.Rows[0].Cells[3].Value.ToString());
                } catch {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.pdf1;
                    // "This document has been password-protected.";
                    return false;
                }
                // オーナーパスワードが掛っているかチェック
                if (pdfReader.IsEncrypted()) {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.pdf2;
                    // "This document has been password-protected.";
                    return false;
                }
                pdfDoc = new iTextSharp.text.Document(pdfReader.GetPageSize(1));
                os = new FileStream(tmpFilePath, FileMode.OpenOrCreate);
                pdfCopy = new PdfCopy(pdfDoc, os);
                // 出力ファイルにパスワード設定
                // rp:ユーザーパスワード
                // wp:オーナーパスワード（空の場合はユーザーパスワードと同じ値を設定）

                pdfCopy.Open();
                if (r_password.Length == 0)
                    rp = null;
                else
                    rp = r_password;
                if (w_password.Length == 0) {
                    wp = r_password;
                    pdfCopy.SetEncryption(
                        PdfCopy.STRENGTH128BITS, rp, wp,
                        PdfCopy.markAll);
                } else {
                    wp = w_password;
                    // AllowPrinting 	印刷
                    // AllowCopy 	内容のコピーと抽出
                    // AllowModifyContents 	文書の変更
                    // AllowModifyAnnotations 	注釈の入力
                    // AllowFillIn 	フォーム・フィールドの入力と署名
                    // AllowScreenReaders 	アクセシビリティのための内容抽出
                    // AllowAssembly 	文書アセンブリ
                    pdfCopy.SetEncryption(
                        PdfCopy.STRENGTH128BITS, rp, wp,
                        PdfCopy.AllowScreenReaders | PdfCopy.AllowPrinting);
                }

                try {
                    // 出力ファイルDocumentを開く
                    pdfDoc.Open();
                    // アップロードPDFファイルの内容を出力ファイルに書き込む
                    pdfCopy.AddDocument(pdfReader);
                    // 出力ファイルDocumentを閉じる
                    pdfDoc.Close();
                    pdfCopy.Close();
                    os.Close();
                    pdfReader.Close();
                    // オリジナルファイルと一時ファイルを置き換える
                    File.Delete(dataGridView1.Rows[0].Cells[3].Value.ToString());
                    File.Move(tmpFilePath, dataGridView1.Rows[0].Cells[3].Value.ToString());
                } catch (Exception eX) {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                    // "Saving failed." + eX.Message;
                    return false;
                }
            }
            return true;
        }

        private Boolean unlockPDF(String r_password, String w_password) {
            Boolean NOPASS = true;

            ////////////////////////////////////////////////////
            // PDF解錠
            ////////////////////////////////////////////////////
            if (avaPDF) {
                // 一時ファイル取得
                String tmpFilePath = Path.GetTempFileName();
                Boolean isRP = false;
                Boolean isWP = false;

                // パスワードなしで読み込めるかチェック
                try {
                    pdfReader = new PdfReader(dataGridView1.Rows[0].Cells[3].Value.ToString());
                    isRP = false; // ユーザーパスワードなし
                                  // オーナーパスワードが掛っているかチェック
                    isWP = (pdfReader.IsEncrypted()) ? true : false;
                    NOPASS = !(isRP || isWP);
                    pdfReader.Close();
                    pdfReader.Dispose();
                } catch {
                    isRP = true;
                    NOPASS = false;
                }
                if (NOPASS) {
                    // パスワードがかかっていない
                    label2.Text = WindowsFormsApplication1.Properties.Resources.pdf2;
                    //"This document is not applied password.";
                    pdfReader.Close();
                    pdfReader.Dispose();
                    return false;
                }
                if (isRP && (r_password.Length == 0)) {
                    // ユーザーパスワードが掛っているが、入力されていない
                    label2.Text = WindowsFormsApplication1.Properties.Resources.pdf3;
                    // "This document has been user password-protected.";
                    return false;
                }
                if (isWP && (w_password.Length == 0)) {
                    // オーナーパスワードが掛っているが、入力されていない
                    label2.Text = WindowsFormsApplication1.Properties.Resources.pdf4;
                    //"This document has been owner password-protected.";
                    return false;
                }

                String rp = (r_password.Length == 0) ? null : r_password;
                String wp = (w_password.Length == 0) ? r_password : w_password;

                try {
                    pdfReader = new PdfReader(dataGridView1.Rows[0].Cells[3].Value.ToString(), (byte[])System.Text.Encoding.ASCII.GetBytes(wp));
                } catch {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.message2;
                    // "Password is incorrect.";
                    return false;
                }


                try {
                    pdfDoc = new iTextSharp.text.Document(pdfReader.GetPageSize(1));
                    os = new FileStream(tmpFilePath, FileMode.OpenOrCreate);
                    pdfCopy = new PdfCopy(pdfDoc, os);
                    pdfCopy.Open();

                    pdfDoc.Open();
                    pdfCopy.AddDocument(pdfReader);

                    pdfDoc.Close();
                    pdfCopy.Close();
                    pdfReader.Close();
                    pdfReader.Dispose();
                    // オリジナルファイルと一時ファイルを置き換える                    
                    System.IO.File.Copy(tmpFilePath, dataGridView1.Rows[0].Cells[3].Value.ToString(), true);
                    System.IO.File.Delete(tmpFilePath);
                } catch (Exception eX) {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                    // "Saving failed." + eX.Message;
                    return false;
                }
            }
            return true;
        }
    }
}