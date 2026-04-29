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
        private Dictionary<string, string> studentMasterDict = new Dictionary<string, string>();

        private PdfDocument pdfDocument;
        private List<int> pagesToProcess = new List<int>();

        private TableLayoutPanel mainLayout;
        private PictureBox picViewer;
        private Panel pnlInput;
        private DataGridView dgvData;

        private TextBox txtStudentId = new TextBox { Width = 220, Font = new Font("맑은 고딕", 18) };
        private Label lblStudentName = new Label { Text = "대기중...", Font = new Font("맑은 고딕", 14, FontStyle.Bold), ForeColor = Color.Gray, AutoSize = true };

        private TextBox txtReason = new TextBox { Width = 220, Font = new Font("맑은 고딕", 18) };

        private Label lblStatus = new Label { Width = 300, ForeColor = Color.Blue, Font = new Font("맑은 고딕", 11, FontStyle.Bold) };
        private Label lblImageProgress = new Label { Width = 300, ForeColor = Color.DarkGreen, Font = new Font("맑은 고딕", 15, FontStyle.Bold) };
        private Button btnExportExcel = new Button { Width = 220, Height = 50, Text = "최종 엑셀 추출", Font = new Font("맑은 고딕", 12, FontStyle.Bold) };

        public Form1()
        {
            SetupUI();
            LoadStudentMasterData();
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
            this.Text = "학사경고자 설문 초고속 입력기 (연락두절/확인요망 특수키 지원)";
            this.Size = new Size(1600, 900);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;

            mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1 };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60f));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20f));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20f));

            picViewer = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.FromArgb(240, 240, 240), Cursor = Cursors.Hand };
            new ToolTip().SetToolTip(picViewer, "마우스로 클릭하면 이 페이지를 건너뜁니다.");
            picViewer.MouseClick += (s, e) => { SkipCurrentPage(); };

            pnlInput = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };
            int yPos = 40;

            pnlInput.Controls.Add(lblImageProgress); lblImageProgress.Location = new Point(20, yPos); yPos += 60;
            pnlInput.Controls.Add(new Label { Text = "★ 불필요한 페이지는 [ESC] 스킵", AutoSize = true, ForeColor = Color.Red, Font = new Font("맑은 고딕", 10), Location = new Point(20, yPos) }); yPos += 40;

            pnlInput.Controls.Add(new Label { Text = "1. 학번 (입력 후 Enter):", AutoSize = true, Font = new Font("맑은 고딕", 12), Location = new Point(20, yPos) });
            yPos += 35;

            pnlInput.Controls.Add(txtStudentId);
            txtStudentId.Location = new Point(20, yPos);
            txtStudentId.KeyDown += TxtStudentId_KeyDown;
            yPos += 45;

            pnlInput.Controls.Add(lblStudentName);
            lblStudentName.Location = new Point(20, yPos);
            lblStudentName.MaximumSize = new Size(250, 0);
            yPos += 40;

            // ★ 라벨 안내 문구 수정 (9와 0에 대한 안내 추가)
            pnlInput.Controls.Add(new Label { Text = "2. 선택 번호 (9=연락두절, 0=확인요망):", AutoSize = true, Font = new Font("맑은 고딕", 11, FontStyle.Bold), Location = new Point(20, yPos) });
            yPos += 35;

            pnlInput.Controls.Add(txtReason);
            txtReason.Location = new Point(20, yPos);
            txtReason.KeyDown += TxtReason_KeyDown;
            txtReason.KeyPress += TxtReason_KeyPress;
            yPos += 70;

            pnlInput.Controls.Add(lblStatus); lblStatus.Location = new Point(20, yPos); yPos += 60;
            pnlInput.Controls.Add(btnExportExcel); btnExportExcel.Location = new Point(20, yPos);
            btnExportExcel.Click += BtnExportExcel_Click;

            // ★ 그리드에 "기타 의견" 컬럼 추가
            dgvData = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, SelectionMode = DataGridViewSelectionMode.FullRowSelect, RowHeadersVisible = false };
            dgvData.Columns.Add("Id", "학번");
            dgvData.Columns.Add("Name", "이름");
            dgvData.Columns.Add("Checks", "선택 번호");
            dgvData.Columns.Add("Other", "기타 특이사항");

            mainLayout.Controls.Add(picViewer, 0, 0);
            mainLayout.Controls.Add(pnlInput, 1, 0);
            mainLayout.Controls.Add(dgvData, 2, 0);

            this.Controls.Add(mainLayout);
        }

        private void LoadStudentMasterData()
        {
            MessageBox.Show("먼저 [학생 명단 CSV 파일]을 선택해주세요.", "명단 파일 선택 (1/2)");
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV Files|*.csv" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var lines = File.ReadAllLines(ofd.FileName, Encoding.Default);
                        foreach (var line in lines)
                        {
                            var parts = line.Split(',');
                            if (parts.Length > 3)
                            {
                                string id = parts[2].Trim();
                                string name = parts[3].Trim();
                                if (id.Length >= 8 && id.All(char.IsDigit)) studentMasterDict[id] = name;
                            }
                        }
                    }
                    catch (Exception ex) { MessageBox.Show("명단 로드 오류: " + ex.Message); }
                }
            }
        }

        private void LoadPdfOnStartup()
        {
            MessageBox.Show("작업하실 [설문지 PDF 파일]을 선택해주세요.", "PDF 열기 (2/2)");
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
            catch { }
        }

        private Image RenderPdfPage(int pageIndex)
        {
            int dpi = 150;
            var size = pdfDocument.PageSizes[pageIndex];
            int width = (int)(size.Width * dpi / 72.0);
            int height = (int)(size.Height * dpi / 72.0);
            return pdfDocument.Render(pageIndex, width, height, dpi, dpi, false);
        }

        private void SkipCurrentPage()
        {
            try { if (pagesToProcess.Count > 0) { pagesToProcess.RemoveAt(0); ShowCurrentPage(); } } catch { }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape) { SkipCurrentPage(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void TxtStudentId_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                string id = txtStudentId.Text.Trim();

                if (studentMasterDict.Count > 0)
                {
                    if (studentMasterDict.ContainsKey(id))
                    {
                        lblStudentName.Text = $"확인된 성명: {studentMasterDict[id]}";
                        lblStudentName.ForeColor = Color.DarkBlue;
                        txtReason.Focus();
                    }
                    else
                    {
                        lblStudentName.Text = "⚠️ 명단에 없는 학번입니다!";
                        lblStudentName.ForeColor = Color.Red;
                        txtStudentId.SelectAll();
                    }
                }
                else txtReason.Focus();
            }
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
            List<string> otherTexts = new List<string>(); // ★ 추가된 기타 의견 리스트

            // ★ 핵심 로직: 문자열을 한 글자씩 검사
            foreach (char c in inputStr)
            {
                if (c >= '1' && c <= '8')
                {
                    int num = c - '0';
                    data.CheckedItems[num - 1] = true;
                    if (!validNumbers.Contains(num)) validNumbers.Add(num);
                }
                else if (c == '9') // 9를 입력하면
                {
                    if (!otherTexts.Contains("연락두절")) otherTexts.Add("연락두절");
                }
                else if (c == '0') // 0을 입력하면
                {
                    if (!otherTexts.Contains("확인요망")) otherTexts.Add("확인요망");
                }
            }
            validNumbers.Sort();

            // 데이터 클래스에 기타 내용 조합해서 넣기 (예: "연락두절, 확인요망")
            data.OtherText = string.Join(", ", otherTexts);
            surveyDict[studentId] = data;

            string checkStr = validNumbers.Count > 0 ? string.Join(", ", validNumbers) : "없음";
            string studentName = studentMasterDict.ContainsKey(studentId) ? studentMasterDict[studentId] : "알수없음";

            // CSV 백업에도 기타 내용을 포함
            using (StreamWriter sw = new StreamWriter(csvFilePath, true, Encoding.UTF8))
            {
                sw.WriteLine($"{data.StudentId},\"{checkStr}\",\"{data.OtherText}\"");
            }

            // 표에도 추가
            dgvData.Rows.Insert(0, studentId, studentName, checkStr, data.OtherText);
            dgvData.Rows[0].Selected = true;
            lblStatus.Text = $"[{studentName}] 데이터 저장 완료!";

            txtStudentId.Clear();
            txtReason.Clear();
            lblStudentName.Text = "대기중...";
            lblStudentName.ForeColor = Color.Gray;
            txtStudentId.Focus();

            SkipCurrentPage();
        }

        private void BtnExportExcel_Click(object sender, EventArgs e)
        {
            if (surveyDict.Count == 0) { MessageBox.Show("데이터가 없습니다."); return; }
            OpenFileDialog openFileDialog = new OpenFileDialog { Title = "엑셀 템플릿 선택", Filter = "Excel Files|*.xlsx" };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string templatePath = openFileDialog.FileName;
                string outputPath = Path.Combine(Path.GetDirectoryName(templatePath), "결과_최종.xlsx");

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
                                worksheet.Cell(row, 9).Value = "o"; // 제출 여부 체크

                                // 1~8번 문항 O 표시
                                for (int i = 0; i < 8; i++)
                                {
                                    if (data.CheckedItems[i]) worksheet.Cell(row, 11 + i).Value = "O";
                                }

                                // ★ 9나 0이 입력되어 OtherText가 생겼다면 S열(19번칸)에 작성
                                if (!string.IsNullOrEmpty(data.OtherText))
                                {
                                    worksheet.Cell(row, 19).Value = data.OtherText;
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
        // ★ 삭제했던 기타 칸(OtherText)을 다시 부활시킴
        public string OtherText { get; set; } = "";
    }
}