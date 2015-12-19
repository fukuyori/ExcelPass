#define EXCEL
#define WORD
#define POWERPOINT
#define PDF

using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;
using iTextSharp.text.pdf;

namespace WindowsFormsApplication1 {
    public partial class Form1 : Form {

#if (EXCEL)
        Microsoft.Office.Interop.Excel.Application xlApp;
        Microsoft.Office.Interop.Excel.Workbooks xlBooks;
        Microsoft.Office.Interop.Excel.Workbook xlBook;
#endif
#if (WORD)
        Microsoft.Office.Interop.Word.Application docApp;
        Microsoft.Office.Interop.Word.Documents docDocs;
        Microsoft.Office.Interop.Word.Document docDoc;
#endif
#if (POWERPOINT)
        Microsoft.Office.Interop.PowerPoint.Application pptApp;
        Microsoft.Office.Interop.PowerPoint.Presentations pptPres;
        Microsoft.Office.Interop.PowerPoint.Presentation pptPre;
#endif
#if (PDF)
        iTextSharp.text.pdf.PdfReader pdfReader;
        iTextSharp.text.pdf.PdfCopy pdfCopy;
        iTextSharp.text.Document pdfDoc;
        FileStream os;
#endif

        Boolean avaExcel = false;
        Boolean avaWord = false;
        Boolean avaPoawrPoint = false;
        Boolean avaPdf = false;
        Boolean isExcel = false;
        Boolean isWord = false;
        Boolean isPowerPoint = false;
        Boolean isPdf = false;

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
#if (EXCEL)
            if (Type.GetTypeFromProgID("Excel.Application") != null) {
                avaExcel = true;
                Pic_Excel.Visible = true;
            }
#endif
#if (WORD)
            if (Type.GetTypeFromProgID("Word.Application") != null) {
                avaWord = true;
                Pic_Word.Visible = true;
            }
#endif
#if (POWERPOINT)
            if (Type.GetTypeFromProgID("PowerPoint.Application") != null) {
                avaPoawrPoint = true;
                Pic_PowerPoint.Visible = true;
            }
#endif
#if (PDF)
            avaPdf = true;
            Pic_Pdf.Visible = true;
#endif

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

        //
        // 施錠開始
        //
        private void button1_Click(object sender, EventArgs e) {
            // 必要なオブジェクトのチェック
            checkType();

            // オブジェクトの割り付け
#if (EXCEL)
            if (avaExcel && isExcel) {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlBooks = xlApp.Workbooks;
            }
#endif
#if (WORD)
            if (avaWord && isWord) {
                docApp = new Microsoft.Office.Interop.Word.Application();
                docDocs = docApp.Documents;
            }
#endif
#if (POWERPOINT)
            if (avaPoawrPoint && isPowerPoint) {
                pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
                pptPres = pptApp.Presentations;
            }
#endif
            procLock();

            // オブジェクトの開放
#if (EXCEL)
            if (avaExcel && isExcel) {
                if (xlBook != null)
                    xlBook = null;
                xlApp.Quit();
            }
#endif
#if (WORD)
            if (avaWord && isWord) {
                if (docDocs != null)
                    docDocs = null;
                ((Microsoft.Office.Interop.Word._Application)docApp).Quit();
                ;
            }
#endif
#if (POWERPOINT)
            if (avaPoawrPoint && isPowerPoint) {
                if (pptPres != null)
                    pptPres = null;
                pptApp.Quit();
            }
#endif
#if (PDF)
            if (avaPdf && isPdf) {
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
#endif
            // ガベージコレクション
            GC.Collect();
        }

        // ファイルの種類をチェック
        private void checkType() {
            isExcel = false;
            isWord = false;
            isPowerPoint = false;
            isPdf = false;

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
                        isPdf = true;
                        break;
                }
            }
        }

        // 施錠処理
        private void procLock() {
            // パスワード設定
            string r_password = textBox1.Text;
            string w_password = textBox2.Text;
            string rp, wp; // パスワード変換用
            string tmpFilePath;

            label2.Text = "";

            if (r_password.Length > 0 || w_password.Length > 0) {

                while (dataGridView1.Rows.Count > 0) {
                    dataGridView1.Rows[0].Cells[0].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[0].Cells[1].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[0].Cells[2].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[0].Cells[3].Style.ForeColor = Color.Red;
#if (EXCEL)
                    ////////////////////////////////////////////////////////
                    // Excel施錠
                    ////////////////////////////////////////////////////////
                    if (avaExcel & dataGridView1.Rows[0].Cells[2].Value.ToString() == "Excel") {
                        // 読み込みパスワードが設定されているかチェック
                        try {
                            xlBook = xlBooks.Open(dataGridView1.Rows[0].Cells[3].Value.ToString(),
                                Type.Missing, true, Type.Missing, "");
                        } catch {
                            label2.Text = WindowsFormsApplication1.Properties.Resources.excel1;
                            //"This workbook has been read password-protected.";
                            // xlBooks.Close();
                            return;
                        }

                        // 書き込みパスワードが設定されているかチェック
                        try {
                            xlBook = xlBooks.Open(dataGridView1.Rows[0].Cells[3].Value.ToString(),
                                Type.Missing, Type.Missing, Type.Missing, "", "");
                        } catch {
                            label2.Text = WindowsFormsApplication1.Properties.Resources.excel2;
                            //"This workbook has been write password-protected.";
                            xlBooks.Close();
                            return;
                        }
                        if (textBox1.Text.Length > 0)
                            xlBook.Password = r_password;
                        if (textBox2.Text.Length > 0)
                            xlBook.WritePassword = w_password;
                        xlBook.CheckCompatibility = false;
                        try {
                            xlBook.Save();
                        } catch (Exception eX) {
                            label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                            //"Saving failed." + eX.Message;
                        }
                        xlBooks.Close();
                    }
#endif
#if (WORD)
                    //////////////////////////////////////////////////////
                    // Word施錠
                    //////////////////////////////////////////////////////
                    if (avaWord & dataGridView1.Rows[0].Cells[2].Value.ToString() == "Word") {
                        // 読み込みパスワードが設定されているかチェック
                        try {
                            docDoc = docDocs.Open(dataGridView1.Rows[0].Cells[3].Value.ToString(),
                                Type.Missing, true, Type.Missing, " ", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false);
                        } catch {
                            label2.Text = WindowsFormsApplication1.Properties.Resources.word1;
                            //"This document has been read password-protected.";
                            docDoc = null;
                            return;
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
                            return;
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
                        }
                        docDoc = null;
                    }
#endif
#if (POWERPOINT)
                    ////////////////////////////////////////////////////
                    // PowerPoint施錠
                    ////////////////////////////////////////////////////
                    if (avaPoawrPoint & dataGridView1.Rows[0].Cells[2].Value.ToString() == "PowerPoint") {
                        // パスワードが設定されているかチェック
                        try {
                            pptPre = pptPres.Open(dataGridView1.Rows[0].Cells[3].Value.ToString() + ":: :: ",
                                Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoFalse);
                        } catch {
                            label2.Text = WindowsFormsApplication1.Properties.Resources.powerpoint1;
                            // "This presentation has been password-protected.";
                            return;
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
                        }
                        pptPre.Close();
                    }
#endif
#if (PDF)
                    ////////////////////////////////////////////////////
                    // PDF施錠
                    ////////////////////////////////////////////////////
                    if (avaPdf & dataGridView1.Rows[0].Cells[2].Value.ToString() == "PDF") {
                        // 一時ファイル取得
                        tmpFilePath = Path.GetTempFileName();

                        // パスワードなしで読み込み可能かチェック
                        try {
                            pdfReader = new PdfReader(dataGridView1.Rows[0].Cells[3].Value.ToString());
                        } catch {
                            label2.Text = WindowsFormsApplication1.Properties.Resources.pdf1;
                            // "This document has been password-protected.";
                            return;
                        }
                        // オーナーパスワードが掛っているかチェック
                        if (pdfReader.IsEncrypted()) {
                            label2.Text = WindowsFormsApplication1.Properties.Resources.pdf2;
                            // "This document has been password-protected.";
                            return;
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
                    }
#endif
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
#if (EXCEL)
            if (avaExcel && isExcel) {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlBooks = xlApp.Workbooks;
            }
#endif
#if (WORD)
            if (avaWord && isWord) {
                docApp = new Microsoft.Office.Interop.Word.Application();
                docDocs = docApp.Documents;
            }
#endif
#if (POWERPOINT)
            if (avaPoawrPoint && isPowerPoint) {
                pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
                pptPres = pptApp.Presentations;
            }
#endif

            procUnlock();

#if (EXCEL)
            if (avaExcel && isExcel) {
                if (xlBook != null)
                    xlBook = null;
                xlApp.Quit();
            }
#endif
#if (WORD)
            if (avaWord && isWord) {
                if (docDocs != null)
                    docDocs = null;
                ((Microsoft.Office.Interop.Word._Application)docApp).Quit();
                ;
            }
#endif
#if (POWERPOINT)
            if (avaPoawrPoint && isPowerPoint) {
                if (pptPres != null)
                    pptPres = null;
                pptApp.Quit();
            }
#endif
#if (PDF)
            if (avaPdf && isPdf) {
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
#endif
            GC.Collect();
        }

        // 解錠処理
        private void procUnlock() {
            // パスワード解除
            string r_password = textBox1.Text;
            string w_password = textBox2.Text;
            string rp = null, wp = null;
            string tmpFilePath;
            Boolean NOPASS = true;
            Boolean isRP, isWP;

            label2.Text = "";

            if (r_password.Length > 0 || w_password.Length > 0) {
                while (dataGridView1.Rows.Count > 0) {
                    dataGridView1.Rows[0].Cells[0].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[0].Cells[1].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[0].Cells[2].Style.ForeColor = Color.Red;
                    dataGridView1.Rows[0].Cells[3].Style.ForeColor = Color.Red;
#if (EXCEL)
                    /////////////////////////////////////////////////////////////
                    // Excel解錠
                    /////////////////////////////////////////////////////////////
                    if (avaExcel & dataGridView1.Rows[0].Cells[2].Value.ToString() == "Excel") {
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
                            return;
                        } else {
                            try {
                                xlBook = xlBooks.Open(dataGridView1.Rows[0].Cells[3].Value.ToString(),
                                    Type.Missing, Type.Missing, Type.Missing, r_password, w_password);
                            } catch {
                                label2.Text = WindowsFormsApplication1.Properties.Resources.message2;
                                // "Password is incorrect.";
                               //  xlBooks.Close();
                                return;
                            }
                            xlBook.Password = "";
                            xlBook.WritePassword = "";
                            xlBook.CheckCompatibility = false;
                            try {
                                xlBook.Save();
                            } catch (Exception eX) {
                                label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                                // "Saving failed." + eX.Message;
                            }
                            xlBooks.Close();
                        }
                    }
#endif
#if (WORD)
                    /////////////////////////////////////////////////////////
                    // Word解錠
                    /////////////////////////////////////////////////////////
                    if (avaWord & dataGridView1.Rows[0].Cells[2].Value.ToString() == "Word") {
                        // 一時ファイル取得
                        tmpFilePath = Path.GetTempFileName();
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
                            return;
                        } else {
                            // パスワード無しの時は、スペース1つ
                            rp = (r_password.Length == 0) ? " " : r_password;
                            wp = (w_password.Length == 0) ? " " : w_password;
                            try {
                                docDoc = docDocs.Open(dataGridView1.Rows[0].Cells[3].Value.ToString(),
                                    Type.Missing, Type.Missing, Type.Missing, rp, Type.Missing, Type.Missing, wp, Type.Missing, Type.Missing, Type.Missing, false);
                            } catch {
                                label2.Text = WindowsFormsApplication1.Properties.Resources.message2;
                                //"Password is incorrect.";
                                docDoc = null;
                                return;
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
                            }
                            docDoc = null;
                        }
                    }
#endif
#if (POWERPOINT)
                    /////////////////////////////////////////////////////
                    // PwerPoint解錠
                    /////////////////////////////////////////////////////
                    if (avaPoawrPoint & dataGridView1.Rows[0].Cells[2].Value.ToString() == "PowerPoint") {
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
                            return;
                        } else {
                            rp = (r_password.Length == 0) ? ":: " : "::" + r_password;
                            wp = (w_password.Length == 0) ? ":: " : "::" + w_password;

                            try {
                                pptPre = pptPres.Open(dataGridView1.Rows[0].Cells[3].Value.ToString() + rp + wp,
                                    Microsoft.Office.Core.MsoTriState.msoFalse,
                                    Microsoft.Office.Core.MsoTriState.msoFalse,
                                    Microsoft.Office.Core.MsoTriState.msoFalse);
                            } catch {
                                label2.Text = WindowsFormsApplication1.Properties.Resources.message2;
                                //"Password is incorrect.";
                                return;
                            }

                            pptPre.Password = null;
                            pptPre.WritePassword = null;
                            try {
                                pptPre.Save();
                            } catch (Exception eX) {
                                label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                                // "Saving failed." + eX.Message;
                            }
                            pptPre.Close();
                        }
                    }
#endif
#if (PDF)
                    ////////////////////////////////////////////////////
                    // PDF解錠
                    ////////////////////////////////////////////////////
                    if (avaPdf & dataGridView1.Rows[0].Cells[2].Value.ToString() == "PDF") {
                        // 一時ファイル取得
                        tmpFilePath = Path.GetTempFileName();
                        isRP = false;
                        isWP = false;

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
                            return;
                        }
                        if (isRP && (r_password.Length == 0)) {
                            // ユーザーパスワードが掛っているが、入力されていない
                            label2.Text = WindowsFormsApplication1.Properties.Resources.pdf3;
                            // "This document has been user password-protected.";
                            return;
                        }
                        if (isWP && (w_password.Length == 0)) {
                            // オーナーパスワードが掛っているが、入力されていない
                            label2.Text = WindowsFormsApplication1.Properties.Resources.pdf4;
                            //"This document has been owner password-protected.";
                            return;
                        }

                        rp = (r_password.Length == 0) ? null : r_password;
                        wp = (w_password.Length == 0) ? r_password : w_password;

                        try {
                            pdfReader = new PdfReader(dataGridView1.Rows[0].Cells[3].Value.ToString(), (byte[])System.Text.Encoding.ASCII.GetBytes(wp));
                        } catch {
                            label2.Text = WindowsFormsApplication1.Properties.Resources.message2;
                            // "Password is incorrect.";
                            return;
                        }

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
                        try {
                            System.IO.File.Copy(tmpFilePath,
                                dataGridView1.Rows[0].Cells[3].Value.ToString(), true);
                            System.IO.File.Delete(tmpFilePath);
                        } catch (Exception eX) {
                            label2.Text = WindowsFormsApplication1.Properties.Resources.error1 + eX.Message;
                            // "Saving failed." + eX.Message;
                        }
                    }
#endif
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
            int idx;
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
#if (EXCEL)
            // Excel
            if (avaExcel & (filePath.Substring(filePath.Length - 3) == "xls" | filePath.Substring(filePath.Length - 4, 3) == "xls")) {
                // DataGridへExcel表追加
                dataGridView1.Rows.Add();
                idx = dataGridView1.Rows.Count - 1;
                dataGridView1.Rows[idx].Cells[1].Value = Path.GetFileName(filePath);
                dataGridView1.Rows[idx].Cells[2].Value = "Excel";
                dataGridView1.Rows[idx].Cells[3].Value = filePath;
            }
#endif
#if (WORD)
            // Word
            if (avaWord & (filePath.Substring(filePath.Length - 3) == "doc" | filePath.Substring(filePath.Length - 4, 3) == "doc")) {
                // DataGridへWord追加
                dataGridView1.Rows.Add();
                idx = dataGridView1.Rows.Count - 1;
                dataGridView1.Rows[idx].Cells[1].Value = Path.GetFileName(filePath);
                dataGridView1.Rows[idx].Cells[2].Value = "Word";
                dataGridView1.Rows[idx].Cells[3].Value = filePath;
            }
#endif
#if (POWERPOINT)
            // PowerPoint
            if (avaPoawrPoint & (filePath.Substring(filePath.Length - 3) == "ppt" | filePath.Substring(filePath.Length - 4, 3) == "ppt")) {
                // DataGridへPowerPoint追加
                dataGridView1.Rows.Add();
                idx = dataGridView1.Rows.Count - 1;
                dataGridView1.Rows[idx].Cells[1].Value = Path.GetFileName(filePath);
                dataGridView1.Rows[idx].Cells[2].Value = "PowerPoint";
                dataGridView1.Rows[idx].Cells[3].Value = filePath;
            }
#endif
#if (PDF)
            // PDF
            if (avaPdf & (filePath.Substring(filePath.Length - 3) == "pdf")) {
                // DataGridへPowerPoint追加
                dataGridView1.Rows.Add();
                idx = dataGridView1.Rows.Count - 1;
                dataGridView1.Rows[idx].Cells[1].Value = Path.GetFileName(filePath);
                dataGridView1.Rows[idx].Cells[2].Value = "PDF";
                dataGridView1.Rows[idx].Cells[3].Value = filePath;
            }
#endif
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