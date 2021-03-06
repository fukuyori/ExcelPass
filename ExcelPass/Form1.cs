﻿using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Collections;

namespace WindowsFormsApplication1 {
    public partial class Form1 : Form {

        Boolean avaExcel = false;
        Boolean avaWord = false;
        Boolean avaPoawrPoint = false;
        Boolean avaPDF = false;
        Boolean avaZip = false;
        Boolean isExcel = false;
        Boolean isWord = false;
        Boolean isPowerPoint = false;
        Boolean isPDF = false;
        Boolean isZip = false;

        public Form1() {
            InitializeComponent();
        }

        //
        // 開始処理
        //
        private void Form1_Load(object sender, EventArgs e) {
            // ドラッグドロップを受け付ける
            dataGridView1.AllowDrop = true;

            // Excel, Word, PowerPointがインストールされているかチェック
            checkExcel();
            checkWord();
            checkPowerPoint();
            checkPDF();
            checkZip();

            this.ActiveControl = textBox1;

            // アプリケーションアイコンへのドラッグ＆ドロップ
            string[] files = System.Environment.GetCommandLineArgs();
            if (files.Length > 1) {
                for (int i = 1; i < files.Length; i++) {
                    if (System.IO.File.Exists(files[i]) == true)
                        addDataGridView(files[i]);
                    else if (System.IO.Directory.Exists(files[i]) == true)
                        // フォルダー内のファイルをすべて
                        foreach (string stFilePath in System.IO.Directory.GetFiles(files[i], "*", System.IO.SearchOption.AllDirectories))
                            if ((System.IO.File.GetAttributes(stFilePath) & System.IO.FileAttributes.Hidden) != System.IO.FileAttributes.Hidden)
                                addDataGridView(stFilePath);
                }
            }
        }

        //
        // 選択したファイルをdataGridViewから消去
        //
        private void button2_Click(object sender, EventArgs e) {
            deleteDataGridView();
        }

        // 選択したデータを消去
        private void deleteDataGridView() {
            // 選択されたCellをRow選択に
            foreach (DataGridViewCell c in dataGridView1.SelectedCells) {
                dataGridView1.Rows[c.RowIndex].Selected = true;
            }
            // Row選択されたアイテムの消去
            while (dataGridView1.SelectedRows.Count != 0) {
                dataGridView1.Rows.Remove(dataGridView1.SelectedRows[0]);
            }
        }

        // 全部消去
        private void deleteAllDataGridView() {
            while (dataGridView1.Rows.Count != 0) {
                dataGridView1.Rows.Remove(dataGridView1.Rows[0]);
            }
        }

        /// 指定されたファイルがロックされているかどうかを返します。
        private bool IsFileLocked(string path) {
            FileStream stream = null;

            try {
                stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            } catch {
                return true;
            } finally {
                if (stream != null) {
                    stream.Close();
                }
            }

            return false;
        }

        //
        // 施錠開始
        //
        private void button1_Click(object sender, EventArgs e) {
            // 必要なオブジェクトのチェック
            checkType();

            // オブジェクトの割り付け
            if (isExcel)
                allocExcel();
            if (isWord)
                allocWord();
            if (isPowerPoint)
                allocPowerPoint();
            if (isPDF)
                allocPDF();
            if (isZip)
                allocZip();

            // 施錠処理
            procLock();

            // オブジェクトの開放
            if (isExcel)
                freeExcel();
            if (isWord)
                freeWord();
            if (isPowerPoint)
                freePowerPoint();
            if (isPDF)
                freePDF();
            if (isZip)
                freeZip();

            // ガベージコレクション
            GC.Collect();
        }

        // ファイルの種類をチェック
        private void checkType() {
            isExcel = false;
            isWord = false;
            isPowerPoint = false;
            isPDF = false;

            for (int i = 0; i < dataGridView1.Rows.Count; i++) {
                switch (dataGridView1.Rows[i].Cells[2].Value.ToString()) {
                    case "Excel":
                        isExcel = true;
                        break;
                    case "Word":
                        isWord = true;
                        break;
                    case "PowerPoint":
                        isPowerPoint = true;
                        break;
                    case "PDF":
                        isPDF = true;
                        break;
                    case "ZIP":
                        isZip = true;
                        break;
                }
            }
        }


        // 施錠処理
        private void procLock() {
            // パスワード設定
            string r_password = textBox1.Text;
            string w_password = textBox2.Text;

            label2.Text = "";

            if (r_password.Length > 0 || w_password.Length > 0) {

                while (dataGridView1.Rows.Count > 0) {
                    dataGridView1.Rows[0].Cells[0].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[0].Cells[1].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[0].Cells[2].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[0].Cells[3].Style.ForeColor = Color.Red;

                    if (IsFileLocked(dataGridView1.Rows[0].Cells[3].Value.ToString())) {
                        label2.Text = WindowsFormsApplication1.Properties.Resources.error2;
                        return;
                    }

                    switch (dataGridView1.Rows[0].Cells[2].Value.ToString()) {
                        case "Excel":
                            if (!lockExcel(r_password, w_password))
                                return;
                            break;
                        case "Word":
                            if (!lockWord(r_password, w_password))
                                return;
                            break;
                        case "PowerPoint":
                            if (!lockPowerPoint(r_password, w_password)) 
                                return;
                            break;
                        case "PDF":
                            if (!lockPDF(r_password, w_password))
                                return;
                            break;
                        case "ZIP":
                            if (!lockZip(r_password, w_password))
                                return;
                            break;
                    }

                    dataGridView1.Rows.RemoveAt(0);
                }
                // 処理完了
                label2.ForeColor = Color.Black;
                label2.BackColor = SystemColors.Control;
                label2.Text = WindowsFormsApplication1.Properties.Resources.done;

            } else {
                label2.Text = WindowsFormsApplication1.Properties.Resources.message1;
                //"Please input password.";
            }
        }

        //
        // 解錠開始
        //
        private void button3_Click(object sender, EventArgs e) {
            // 必要なオブジェクトのチェック
            checkType();

            // オブジェクトの割り付け
            if (isExcel)
                allocExcel();
            if (isWord)
                allocWord();
            if (isPowerPoint)
                allocPowerPoint();
            if (isPDF)
                allocPDF();
            if (isZip)
                allocZip();

            // 解錠処理
            procUnlock();

            // オブジェクトの開放
            if (isExcel)
                freeExcel();
            if (isWord)
                freeWord();
            if (isPowerPoint)
                freePowerPoint();
            if (isPDF)
                freePDF();
            if (isZip)
                freeZip();

            GC.Collect();
        }

        // 解錠処理
        private void procUnlock() {
            // パスワード解除
            string r_password = textBox1.Text;
            string w_password = textBox2.Text;

            label2.Text = "";

            if (r_password.Length > 0 || w_password.Length > 0) {
                while (dataGridView1.Rows.Count > 0) {
                    dataGridView1.Rows[0].Cells[0].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[0].Cells[1].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[0].Cells[2].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[0].Cells[3].Style.ForeColor = Color.Red;

                    if (IsFileLocked(dataGridView1.Rows[0].Cells[3].Value.ToString())) {
                        label2.Text = WindowsFormsApplication1.Properties.Resources.error2;
                        return;
                    }

                    switch (dataGridView1.Rows[0].Cells[2].Value.ToString()) {
                        case "Excel":
                            if (!unlockExcel(r_password, w_password))
                                return;
                            break;
                        case "Word":
                            if (!unlockWord(r_password, w_password))
                                return;
                            break;
                        case "PowerPoint":
                            if (!unlockPowerPoint(r_password, w_password))
                                return;
                            break;
                        case "PDF":
                            if (!unlockPDF(r_password, w_password))
                                return;
                            break;
                        case "ZIP":
                            if (!unlockZip(r_password, w_password))
                                return;
                            break;
                    }

                    dataGridView1.Rows.RemoveAt(0);
                }
                // 処理完了
                label2.ForeColor = Color.Black;
                label2.BackColor = SystemColors.Control;
                label2.Text = WindowsFormsApplication1.Properties.Resources.done;

            } else {
                label2.Text = WindowsFormsApplication1.Properties.Resources.message1;
                // "Please input password.";
            }
        }

        //
        // ファイルがドロップされた時の処理
        //
        private void dataGridView1_DragDrop(object sender, DragEventArgs e) {
            System.IO.FileAttributes fattr;

            // ファイルが渡されていなければ、何もしない
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return;

            foreach (var filePath in (string[])e.Data.GetData(DataFormats.FileDrop)) {
                if (System.IO.File.Exists(filePath) == true)
                    addDataGridView(filePath);
                else if (System.IO.Directory.Exists(filePath) == true)
                    // フォルダー内のファイルをすべて
                    foreach (string stFilePath in System.IO.Directory.GetFiles(filePath, "*", System.IO.SearchOption.AllDirectories))
                        if ((System.IO.File.GetAttributes(stFilePath) & System.IO.FileAttributes.Hidden) != System.IO.FileAttributes.Hidden)
                            addDataGridView(stFilePath);
            }

            if (dataGridView1.Rows.Count > 0) {
                dataGridView1.Rows[0].Cells[0].Style.ForeColor = Color.Black;
                dataGridView1.Rows[0].Cells[1].Style.ForeColor = Color.Black;
                dataGridView1.Rows[0].Cells[2].Style.ForeColor = Color.Black;
                dataGridView1.Rows[0].Cells[3].Style.ForeColor = Color.Black;
            }
            label2.Text = "";
        }

        // ドラッグドロップ時にカーソルの形状を変更
        private void dataGridView1_DragEnter(object sender, DragEventArgs e) {
            e.Effect = DragDropEffects.All;
        }

        // DataGridViewへファイル追加
        private void addDataGridView(string filePath) {
 
            // 同じファイルがあったら、追加しない
            Boolean SAMEFILE = false;
            foreach (DataGridViewRow row in dataGridView1.Rows) {
                if (row.Cells[3].Value.ToString() == filePath) {
                    SAMEFILE = true;
                    break;
                }
            }
            if (SAMEFILE)
                return; // 同じファイルが登録済みなら、追加しない

            switch (filePath.Substring(filePath.Length - 3).ToUpper()) {
                case "XLS":
                    if (isOle2File(filePath))
                        addExcel(filePath);
                    break;
                case "DOC":
                    if (isOle2File(filePath))
                        addWord(filePath);
                    break;
                case "PPT":
                    if (isOle2File(filePath))
                        addPowerPoint(filePath);
                    break;
                case "PDF":
                    if (isPdfFile(filePath))
                        addPDF(filePath);
                    break;
                case "ZIP":
                    if (isZipFile(filePath))
                        addZip(filePath);
                    break;
            }
            // 拡張子４桁のファイル
            switch (filePath.Substring(filePath.Length - 4, 3).ToUpper()) {
                case "XLS":
                    if (isOpenXMLFile(filePath))
                        addExcel(filePath);
                    break;
                case "DOC":
                    if (isOpenXMLFile(filePath))
                        addWord(filePath);
                    break;
                case "PPT":
                    if (isOpenXMLFile(filePath))
                        addPowerPoint(filePath);
                    break;
            }
        }

        // OLE2フォーマットのOffice文書か調べる (doc. xls, ppt)
        private bool isOle2File(string filePath) {
            byte[] sig = new byte[8] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };

            System.IO.FileStream fs = new System.IO.FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read);

            byte[] bs = new byte[sig.Length];

            try {
                fs.Read(bs, 0, bs.Length);
            } catch {
                fs.Close();
                return false;
            }

            fs.Close();

            //2つの配列が等しいか調べる
            return ((IStructuralEquatable)sig).Equals(bs,
                StructuralComparisons.StructuralEqualityComparer);
        }

        // Office2007以降のファイルかチェック (xlsx, docx, pptx)
        private bool isOpenXMLFile(string filePath) {
            byte[] sig1 = new byte[8] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
            byte[] sig2 = new byte[8] { 0x50, 0x4b, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00 };

            System.IO.FileStream fs = new System.IO.FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read);

            byte[] bs = new byte[sig1.Length];

            try {
                fs.Read(bs, 0, bs.Length);
            } catch {
                fs.Close();
                return false;
            }

            fs.Close();

            //2つの配列が等しいか調べる
            return
                ((IStructuralEquatable)sig1).Equals(bs, StructuralComparisons.StructuralEqualityComparer)
                | ((IStructuralEquatable)sig2).Equals(bs, StructuralComparisons.StructuralEqualityComparer);
        }

        // PDFファイルか確認
        private bool isPdfFile(string filePath) {
            byte[] sig = new byte[4] { 0x25, 0x50, 0x44, 0x46 }; //%PDF

            System.IO.FileStream fs = new System.IO.FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read);

            byte[] bs = new byte[sig.Length];

            try {
                fs.Read(bs, 0, bs.Length);
            } catch {
                fs.Close();
                return false;
            }

            fs.Close();

            //2つの配列が等しいか調べる
            return ((IStructuralEquatable)sig).Equals(bs,
                StructuralComparisons.StructuralEqualityComparer);
        }

        // ZIP ファイルかチェック
        private bool isZipFile(string filePath) {
            System.IO.FileStream fs = new System.IO.FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read);

            //ZipInputStreamオブジェクトの作成
            ICSharpCode.SharpZipLib.Zip.ZipInputStream zis =
                new ICSharpCode.SharpZipLib.Zip.ZipInputStream(fs);

            try {
                ICSharpCode.SharpZipLib.Zip.ZipEntry ze;
                while ((ze = zis.GetNextEntry()) != null) {
                    //Console.WriteLine(ze.Name);
                    fs.Close();
                    return true;
                }
            } catch {
                fs.Close();
                return false;
            }

            fs.Close();
            return false;
        }

        // 終了時にはガベージコレクションを実行
        private void Form1_FormClosed(object sender, FormClosedEventArgs e) {
            GC.Collect();
        }

        // バージョン情報を表示
        private void button4_Click(object sender, EventArgs e) {
            About f = new About();
            f.ShowDialog(this);
            f.Dispose();
        }

        // DELETEキーが押されたら、DataGridViewで選択されているファイルを消去
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyData == Keys.Delete) {
                deleteDataGridView();
            }
        }

        //
        // ToolStripMenu 1
        //
        // 消去
        private void clearToolStripMenuItem1_Click(object sender, EventArgs e) {
            textBox1.Clear();
        }
        // コピー
        private void copyToolStripMenuItem1_Click(object sender, EventArgs e) {
            Clipboard.SetDataObject(textBox1.Text, true);
        }
        // ペースト
        private void pasteToolStripMenuItem1_Click(object sender, EventArgs e) {
            string text = Clipboard.GetText();
            if (!string.IsNullOrEmpty(text)) {
                textBox1.Text = text;
            }
        }
        // パスワード生成（15桁）
        private void createPasswordToolStripMenuItem1_Click(object sender, EventArgs e) {
            textBox1.Text = createPassword(15);
        }

        //
        // ToolsStripMenu 2
        //
        // 消去
        private void clearToolStripMenuItem2_Click(object sender, EventArgs e) {
            textBox2.Clear();
        }
        // コピー
        private void copyToolStripMenuItem2_Click(object sender, EventArgs e) {
            Clipboard.SetDataObject(textBox2.Text, true);
        }
        // ペースト
        private void pasteToolStripMenuItem2_Click(object sender, EventArgs e) {
            string text = Clipboard.GetText();
            if (!string.IsNullOrEmpty(text)) {
                textBox2.Text = text;
            }
        }
        // パスワード生成（15桁）
        private void createPasswordToolStripMenuItem2_Click(object sender, EventArgs e) {
            textBox2.Text = createPassword(15);
        }

        // パスワード作成処理
        private static readonly string passwordChars = "23456789abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ"; // lI1や0Oは間違いやすいので除外
        private string createPassword(int length) {
            /// ランダムな文字列を生成する
            string pass = "";
            Random r = new Random();

            for (int i = 0; i < length; i++) {
                int pos = r.Next(passwordChars.Length);
                char c = passwordChars[pos];
                pass = pass + c;
            }
            Clipboard.SetDataObject(pass.ToString(), true);

            return pass.ToString();
        }

        //
        // ToolStripMenu 3
        //
        // 消去
        private void clearToolStripMenuItem_Click(object sender, EventArgs e) {
            deleteDataGridView();
        }
        // 全消去
        private void clearAllToolStripMenuItem_Click(object sender, EventArgs e) {
            deleteAllDataGridView();
        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e) {
            label6.Text = dataGridView1.Rows.Count.ToString();
            label2.ForeColor = Color.White;
            label2.BackColor = Color.Red;
        }

        private void dataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e) {
            label6.Text = dataGridView1.Rows.Count.ToString();
            label2.ForeColor = Color.White;
            label2.BackColor = Color.Red;
        }
    }
}