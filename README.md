# CSharp-DatagridView-Export

```sh
## ---- > Button 1 (for  SafeFileDialog)

this.Cursor = Cursors.AppStarting;

            /* 1 - excel var mi ? */
            Type officeType = Type.GetTypeFromProgID("Excel");
            if (officeType == null)
            {
                System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
                //default masa üstü gelsin
                saveDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                //sadece excel dosyası kaydedilecek filtresi
                saveDlg.Filter = "Excel dosyası (*.xls)|*.xlsx";
                saveDlg.FilterIndex = 0;
                saveDlg.RestoreDirectory = true;
                //açılan pencere ismi
                saveDlg.Title = "Excel'e Aktar";
                //açılan pencerede default dosya ismi
                saveDlg.FileName = "TamirKayitlarim-" + tarih;

                if (saveDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    //dosya yolunu alıyoruz
                    path = saveDlg.FileName;

                    if (backgroundWorker1.IsBusy == false)
                    {
                        backgroundWorker1.RunWorkerAsync();
                    }
                }
                else
                {
                    MessageBox.Show("İşlem iptal edildi.", "Bildirim", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Bu bilgisayarda excel uygulaması bulunamadı.", "Bildirim", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            this.Cursor = Cursors.Default; 
            
## ---- > BackgroundWorker 1 (for  async)
            
            pictureBox3.Image = Properties.Resources.icons8_save_as_80;
            label5.Text = "Kayıt Ediliyor";
            pictureBox3.Enabled = false;

            Excel.Application ExcelApp = new Excel.Application();
            ExcelApp.Workbooks.Add();
            ExcelApp.Visible = false; // <-- kullanıcı görmesin diye false yapıyoruz

            /* 2.sayfa *******************/
            ExcelApp.Worksheets[1].name = "Data";
            ExcelApp.Worksheets[1].Activate();

            /* datagrid verilerini 2.sayfaya aktarma yeri */
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Range myRange = (Range)ExcelApp.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }

            StartRow++;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    Range myRange = (Range)ExcelApp.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    myRange.Select();
                }
            }

            ExcelApp.Workbooks[1].SaveCopyAs(path);
            ExcelApp.Workbooks[1].Saved = true;
            ExcelApp.Workbooks.Close();
            ExcelApp.Quit();
```
