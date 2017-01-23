using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

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
                    addDataGridView(files[i]);
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
            } else {
                label2.Text = WindowsFormsApplication1.Properties.Resources.message1;
                // "Please input password.";
            }
        }

        //
        // ファイルがドロップされた時の処理
        //
        private void dataGridView1_DragDrop(object sender, DragEventArgs e) {
            // ファイルが渡されていなければ、何もしない
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return;

            foreach (var filePath in (string[])e.Data.GetData(DataFormats.FileDrop)) {
                addDataGridView(filePath);
            }
            dataGridView1.Rows[0].Cells[0].Style.ForeColor = Color.Black;
            dataGridView1.Rows[0].Cells[1].Style.ForeColor = Color.Black;
            dataGridView1.Rows[0].Cells[2].Style.ForeColor = Color.Black;
            dataGridView1.Rows[0].Cells[3].Style.ForeColor = Color.Black;
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

            // 拡張子３桁のファイル
            switch (filePath.Substring(filePath.Length - 3).ToUpper()) {
                case "XLS":
                    addExcel(filePath);
                    break;
                case "DOC":
                    addWord(filePath);
                    break;
                case "PPT":
                    addPowerPoint(filePath);
                    break;
                case "PDF":
                    addPDF(filePath);
                    break;
                case "ZIP":
                    addZip(filePath);
                    break;
            }
            // 拡張子４桁のファイル
            switch (filePath.Substring(filePath.Length - 4, 3).ToUpper()) {
                case "XLS":
                    addExcel(filePath);
                    break;
                case "DOC":
                    addWord(filePath);
                    break;
                case "PPT":
                    addPowerPoint(filePath);
                    break;
            }
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
        // パスワード生成（10桁）
        private void createPasswordToolStripMenuItem1_Click(object sender, EventArgs e) {
            textBox1.Text = createPassword(10);
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
        // パスワード生成（10桁）
        private void createPasswordToolStripMenuItem2_Click(object sender, EventArgs e) {
            textBox2.Text = createPassword(10);
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
    }
}