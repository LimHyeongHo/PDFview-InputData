using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;
using PdfiumViewer;

namespace SurveyDataEntry
{
    public partial class Form1 : Form
    {
        private Dictionary<string, SurveyData> surveyDict = new Dictionary<string, SurveyData>();
        private string csvFilePath = "survey_backup.csv";

        private PdfDocument pdfDocument;
        private List<int> pagesToProcess = new List<int>();

        private TableLayoutPanel mainLayout;
        private PictureBox picViewer; // 1장 크게 보기용
        private Panel pnlInput;
        private DataGridView dgvData;

        private TextBox txtStudentId = new TextBox { Width = 220, Font = new Font("맑은 고딕", 18) };
        private TextBox txtReason = new TextBox { Width = 220, Font = new Font("맑은 고딕", 18) };

        private Label lblStatus = new Label { Width = 300, ForeColor = Color.Blue, Font = new Font("맑은 고딕", 11, FontStyle.Bold) };
        private Label lblImageProgress = new Label { Width = 300, ForeColor = Color.DarkGreen, Font = new Font("맑은 고딕", 15, FontStyle.Bold) };
        private Button btnExportExcel = new Button { Width = 220, Height = 50, Text = "최종 엑셀 추출", Font = new Font("맑은 고딕", 12, FontStyle.Bold) };

        public Form1()
        {
            SetupUI();
            LoadPdfOnStartup();
        }

        private void Form1_Load(object sender, EventArgs e) { }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            pdfDocument?.Dispose();
            base.OnFormClosed(e);
        }

        private void SetupUI()
        {
            this.Text = "학사경고자 설문 전문 입력 시스템 (1장 크게 보기 모드)";
            this.Size = new Size(1600, 900);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;

            mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1 };
            // 비율 조정: 사진(60%) : 입력(20%) : 확인(20%)
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60f));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20f));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20f));

            // 좌측: 사진 1장 크게 보기
            picViewer = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.FromArgb(240, 240, 240),
                Cursor = Cursors.Hand
            };

            ToolTip toolTip = new ToolTip();
            toolTip.SetToolTip(picViewer, "마우스로 클릭하면 이 페이지를 건너뜁니다.");

            // 마우스 클릭 시 1장 스킵
            picViewer.MouseClick += (s, e) => { SkipCurrentPage(); };

            // 중앙: 입력 패널
            pnlInput = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };
            int yPos = 40;

            pnlInput.Controls.Add(lblImageProgress); lblImageProgress.Location = new Point(20, yPos); yPos += 60;
            pnlInput.Controls.Add(new Label { Text = "★ 팁: 불필요한 페이지는 [ESC]나 [사진 클릭]으로 1장씩 넘기세요.", AutoSize = true, ForeColor = Color.Red, Font = new Font("맑은 고딕", 9), Location = new Point(20, yPos) }); yPos += 40;

            pnlInput.Controls.Add(new Label { Text = "1. 학번 (입력 후 Enter):", AutoSize = true, Font = new Font("맑은 고딕", 12), Location = new Point(20, yPos) }); yPos += 35;
            pnlInput.Controls.Add(txtStudentId); txtStudentId.Location = new Point(20, yPos); yPos += 70;
            txtStudentId.KeyDown += TxtStudentId_KeyDown;

            pnlInput.Controls.Add(new Label { Text = "2. 선택 번호 (예: 124 누르고 Enter):", AutoSize = true, Font = new Font("맑은 고딕", 12), Location = new Point(20, yPos) }); yPos += 35;
            pnlInput.Controls.Add(txtReason); txtReason.Location = new Point(20, yPos); yPos += 70;
            txtReason.KeyDown += TxtReason_KeyDown;
            txtReason.KeyPress += TxtReason_KeyPress;

            pnlInput.Controls.Add(lblStatus); lblStatus.Location = new Point(20, yPos); yPos += 60;
            pnlInput.Controls.Add(btnExportExcel); btnExportExcel.Location = new Point(20, yPos);
            btnExportExcel.Click += BtnExportExcel_Click;

            // 우측: 그리드
            dgvData = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, SelectionMode = DataGridViewSelectionMode.FullRowSelect, RowHeadersVisible = false };
            dgvData.Columns.Add("Id", "학번");
            dgvData.Columns.Add("Checks", "선택 번호");

            mainLayout.Controls.Add(picViewer, 0, 0);
            mainLayout.Controls.Add(pnlInput, 1, 0);
            mainLayout.Controls.Add(dgvData, 2, 0);

            this.Controls.Add(mainLayout);
        }

        private void LoadPdfOnStartup()
        {
            MessageBox.Show("PDF 파일을 선택해주세요.", "PDF 열기");
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "PDF Files|*.pdf" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    pdfDocument = PdfDocument.Load(ofd.FileName);
                    pagesToProcess = Enumerable.Range(0, pdfDocument.PageCount).ToList();
                    ShowCurrentPage();
                }
            }
        }

        private void ShowCurrentPage()
        {
            if (picViewer.Image != null) { picViewer.Image.Dispose(); picViewer.Image = null; }

            if (pagesToProcess.Count == 0)
            {
                lblImageProgress.Text = "모든 작업 완료!";
                picViewer.BackColor = Color.White;
                return;
            }

            try
            {
                int pageIndex = pagesToProcess[0];
                picViewer.Image = RenderPdfPage(pageIndex);
                lblImageProgress.Text = $"현재 페이지: {pageIndex + 1} / {pdfDocument.PageCount}";
            }
            catch (Exception ex)
            {
                MessageBox.Show("이미지 로드 중 오류: " + ex.Message);
            }
        }

        private Image RenderPdfPage(int pageIndex)
        {
            int dpi = 150;
            var size = pdfDocument.PageSizes[pageIndex];
            int width = (int)(size.Width * dpi / 72.0);
            int height = (int)(size.Height * dpi / 72.0);
            return pdfDocument.Render(pageIndex, width, height, dpi, dpi, false);
        }

        // 1장 스킵 공통 로직
        private void SkipCurrentPage()
        {
            try
            {
                if (pagesToProcess.Count > 0)
                {
                    pagesToProcess.RemoveAt(0);
                    ShowCurrentPage();
                }
            }
            catch { }
        }

        // ESC 키로 1장 건너뛰기
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                SkipCurrentPage();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void TxtStudentId_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; txtReason.Focus(); }
        }

        private void TxtReason_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != 13) e.Handled = true;
        }

        private void TxtReason_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; SaveCurrentData(); }
        }

        private void SaveCurrentData()
        {
            string studentId = txtStudentId.Text.Trim();
            if (string.IsNullOrEmpty(studentId)) return;

            var data = new SurveyData { StudentId = studentId };
            string inputStr = txtReason.Text.Trim();
            List<int> validNumbers = new List<int>();

            foreach (char c in inputStr)
            {
                if (c >= '1' && c <= '8')
                {
                    int num = c - '0';
                    data.CheckedItems[num - 1] = true;
                    if (!validNumbers.Contains(num)) validNumbers.Add(num);
                }
            }
            validNumbers.Sort();

            surveyDict[studentId] = data;
            string checkStr = validNumbers.Count > 0 ? string.Join(", ", validNumbers) : "없음";

            using (StreamWriter sw = new StreamWriter(csvFilePath, true, Encoding.UTF8))
            {
                sw.WriteLine($"{data.StudentId},\"{checkStr}\"");
            }

            dgvData.Rows.Insert(0, studentId, checkStr);
            dgvData.Rows[0].Selected = true;
            lblStatus.Text = $"[{studentId}] 저장 완료!";

            txtStudentId.Clear();
            txtReason.Clear();
            txtStudentId.Focus();

            // 저장 후 현재 페이지 1장만 넘기기
            SkipCurrentPage();
        }

        private void BtnExportExcel_Click(object sender, EventArgs e)
        {
            if (surveyDict.Count == 0) { MessageBox.Show("데이터가 없습니다."); return; }
            OpenFileDialog openFileDialog = new OpenFileDialog { Title = "엑셀 템플릿 선택", Filter = "Excel Files|*.xlsx" };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string templatePath = openFileDialog.FileName;
                string outputPath = Path.Combine(Path.GetDirectoryName(templatePath), "수요조사결과_최종.xlsx");

                try
                {
                    using (var workbook = new XLWorkbook(templatePath))
                    {
                        var worksheet = workbook.Worksheet(1);
                        int startRow = 9;
                        int lastRow = worksheet.LastRowUsed().RowNumber();

                        for (int row = startRow; row <= lastRow; row++)
                        {
                            string studentId = worksheet.Cell(row, 3).GetString().Trim();
                            if (surveyDict.ContainsKey(studentId))
                            {
                                var data = surveyDict[studentId];
                                worksheet.Cell(row, 9).Value = "o";
                                for (int i = 0; i < 8; i++)
                                {
                                    if (data.CheckedItems[i]) worksheet.Cell(row, 11 + i).Value = "O";
                                }
                            }
                        }
                        workbook.SaveAs(outputPath);
                    }
                    MessageBox.Show("엑셀 추출 완료!");
                }
                catch (Exception ex) { MessageBox.Show("에러: " + ex.Message); }
            }
        }
    }

    public class SurveyData
    {
        public string StudentId { get; set; }
        public bool[] CheckedItems { get; set; } = new bool[8];
    }
}