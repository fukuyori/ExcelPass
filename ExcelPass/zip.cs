using System;
using System.IO;
using System.Windows.Forms;
using ICSharpCode;

namespace WindowsFormsApplication1 {
    public partial class Form1 : Form {

        ICSharpCode.SharpZipLib.Zip.FastZip myZip;

        private void checkZip() {
            avaZip = true;
            Pic_Zip.Visible = true;
        }

        private void addZip(String filePath) {
            if (avaPDF) {
                // DataGridへZip追加
                dataGridView1.Rows.Add();
                int idx = dataGridView1.Rows.Count - 1;
                dataGridView1.Rows[idx].Cells[1].Value = Path.GetFileName(filePath);
                dataGridView1.Rows[idx].Cells[2].Value = "ZIP";
                dataGridView1.Rows[idx].Cells[3].Value = filePath;
            }
        }

        private void allocZip() {
            if (avaZip) {
                //FastZipオブジェクトの作成
                myZip = new ICSharpCode.SharpZipLib.Zip.FastZip();
                //属性を復元
                myZip.RestoreAttributesOnExtract = true;
                //ファイル日時を復元
                myZip.RestoreDateTimeOnExtract = true;
                //空のフォルダも作成
                myZip.CreateEmptyDirectories = true;
            }
        }

        private void freeZip() {
            if (avaZip) {
                myZip = null;
            }
        }

        private Boolean lockZip(String r_password, String w_password) {
            ////////////////////////////////////////////////////
            // ZIP施錠
            ////////////////////////////////////////////////////
            if (avaZip) {
                if (r_password.Length == 0 || w_password.Length > 0) {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.message3;
                    return false;
                }
                // 一時ファイルの保存フォルダ名
                String tmpFilePath = Path.GetTempPath() + Path.GetRandomFileName();

                // パスワードなしで解凍を行う
                myZip.Password = null;
                try {
                    myZip.ExtractZip(dataGridView1.Rows[0].Cells[3].Value.ToString(), tmpFilePath, null);
                } catch {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.zip1;
                    //"This workbook has been read password-protected.";
                    return false;
                }

                try {
                    // 元のファイルを名前を変更
                    if (File.Exists(dataGridView1.Rows[0].Cells[3].Value.ToString() + "~"))
                        File.Delete(dataGridView1.Rows[0].Cells[3].Value.ToString() + "~");
                    File.Move(dataGridView1.Rows[0].Cells[3].Value.ToString(), dataGridView1.Rows[0].Cells[3].Value.ToString() + "~");
                    // パスワードを付けてファイルを作成
                    myZip.Password = r_password;
                    myZip.CreateZip(dataGridView1.Rows[0].Cells[3].Value.ToString(), tmpFilePath, true, null, null);
                    // 元ファイルを削除
                    File.Delete(dataGridView1.Rows[0].Cells[3].Value.ToString() + "~");
                    // 一時ファイルディレクトリを削除
                    Directory.Delete(tmpFilePath, true);
                } catch {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.error1;
                    Directory.Delete(tmpFilePath, true);
                    return false;
                }
            }
            return true;
        }

        private Boolean unlockZip(String r_password, String w_password) {
            ////////////////////////////////////////////////////
            // ZIP解錠
            ////////////////////////////////////////////////////
            if (avaZip) {
                String tmpFilePath = Path.GetTempPath() + Path.GetRandomFileName();
                Boolean isRP = false;

                // パスワードなしで解凍を行う
                myZip.Password = null;
                try {
                    myZip.ExtractZip(dataGridView1.Rows[0].Cells[3].Value.ToString(), tmpFilePath, null);
                } catch {
                    isRP = true;
                }
                // パスワードがかかっていなければエラー
                if (!isRP) {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.zip2;
                    return false;
                }

                myZip.Password = r_password;
                try {
                    myZip.ExtractZip(dataGridView1.Rows[0].Cells[3].Value.ToString(), tmpFilePath, null);
                } catch {
                    // パスワード間違い
                    label2.Text = WindowsFormsApplication1.Properties.Resources.message2;
                    return false;
                }

                try {
                    // 元のファイルを名前を変更
                    if (File.Exists(dataGridView1.Rows[0].Cells[3].Value.ToString() + "~"))
                        File.Delete(dataGridView1.Rows[0].Cells[3].Value.ToString() + "~");
                    File.Move(dataGridView1.Rows[0].Cells[3].Value.ToString(), dataGridView1.Rows[0].Cells[3].Value.ToString() + "~");
                    // パスワードを付けてファイルを作成
                    myZip.Password = null;
                    myZip.CreateZip(dataGridView1.Rows[0].Cells[3].Value.ToString(), tmpFilePath, true, null, null);
                    // 元ファイルを削除
                    File.Delete(dataGridView1.Rows[0].Cells[3].Value.ToString() + "~");
                    // 一時ファイルディレクトリを削除
                    Directory.Delete(tmpFilePath, true);
                } catch {
                    label2.Text = WindowsFormsApplication1.Properties.Resources.error1;
                    return false;
                }
            }
            return true;
        }
    }
}