using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;
using PdfiumViewer; // PDF 엔진 추가

namespace SurveyDataEntry
{
    public partial class Form1 : Form
    {
        private Dictionary<string, SurveyData> surveyDict = new Dictionary<string, SurveyData>();
        private string csvFilePath = "survey_backup.csv";

        private PdfDocument pdfDocument; // 로드된 PDF 문서 객체
        private List<int> pagesToProcess = new List<int>(); // 아직 입력하지 않은 남은 페이지 번호 대기열

        private TableLayoutPanel mainLayout;
        private TableLayoutPanel imageLayout;
        private PictureBox picLeft;
        private PictureBox picRight;
        private Panel pnlInput;
        private DataGridView dgvData;

        private TextBox txtStudentId = new TextBox { Width = 150, Font = new Font("맑은 고딕", 14) };
        private TextBox[] txtChecks = new TextBox[8];
        private TextBox txtOtherText = new TextBox { Width = 250, Font = new Font("맑은 고딕", 14) };
        private Label lblStatus = new Label { Width = 250, ForeColor = Color.Blue, Font = new Font("맑은 고딕", 10, FontStyle.Bold) };
        private Label lblImageProgress = new Label { Width = 250, ForeColor = Color.DarkGreen, Font = new Font("맑은 고딕", 12, FontStyle.Bold) };
        private Button btnExportExcel = new Button { Width = 150, Height = 40, Text = "최종 엑셀 추출", Font = new Font("맑은 고딕", 10) };

        public Form1()
        {
            SetupUI();
            LoadPdfOnStartup();
        }

        private void Form1_Load(object sender, EventArgs e) { }

        // 폼이 닫힐 때 PDF 엔진 메모리 해제
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            pdfDocument?.Dispose();
            base.OnFormClosed(e);
        }

        private void SetupUI()
        {
            this.Text = "학사경고자 설문 전문 입력 시스템 (PDF 전용 + 클릭 스킵)";
            this.Size = new Size(1600, 900);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;

            mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1 };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));

            imageLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            imageLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            imageLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));

            picLeft = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.LightGray, Cursor = Cursors.Hand };
            picRight = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.DarkGray, Cursor = Cursors.Hand };

            // 툴팁(안내문) 추가
            ToolTip toolTip = new ToolTip();
            toolTip.SetToolTip(picLeft, "클릭하면 이 페이지를 건너뜁니다 (삭제)");
            toolTip.SetToolTip(picRight, "클릭하면 이 페이지를 건너뜁니다 (삭제)");

            // 마우스 클릭 시 해당 이미지만 대기열에서 삭제하고 새로고침
            picLeft.MouseClick += (s, e) => {
                if (pagesToProcess.Count > 0)
                {
                    pagesToProcess.RemoveAt(0);
                    ShowCurrentPages();
                }
            };
            picRight.MouseClick += (s, e) => {
                if (pagesToProcess.Count > 1)
                {
                    pagesToProcess.RemoveAt(1);
                    ShowCurrentPages();
                }
            };

            imageLayout.Controls.Add(picLeft, 0, 0);
            imageLayout.Controls.Add(picRight, 1, 0);

            // 중앙 입력 패널
            pnlInput = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };
            int yPos = 30;

            pnlInput.Controls.Add(lblImageProgress); lblImageProgress.Location = new Point(20, yPos); yPos += 40;

            pnlInput.Controls.Add(new Label { Text = "1. 학번 (입력 후 Enter):", AutoSize = true, Location = new Point(20, yPos) }); yPos += 25;
            pnlInput.Controls.Add(txtStudentId); txtStudentId.Location = new Point(20, yPos); yPos += 50;
            txtStudentId.KeyDown += TxtStudentId_KeyDown;

            pnlInput.Controls.Add(new Label { Text = "2. 체크항목 (1 입력=체크, 빈칸=Enter로 넘김):", AutoSize = true, Location = new Point(20, yPos) }); yPos += 30;

            for (int i = 0; i < 8; i++)
            {
                int col = i % 2;
                int row = i / 2;

                Label lbl = new Label { Text = $"{i + 1}번:", Width = 40, TextAlign = ContentAlignment.MiddleRight };
                lbl.Location = new Point(20 + col * 120, yPos + row * 40);

                txtChecks[i] = new TextBox { Width = 50, Font = new Font("맑은 고딕", 14), TextAlign = HorizontalAlignment.Center };
                txtChecks[i].Location = new Point(65 + col * 120, yPos + row * 40);

                int currentIndex = i;
                txtChecks[i].KeyDown += (s, e) =>
                {
                    if (e.KeyCode == Keys.Enter)
                    {
                        e.SuppressKeyPress = true;
                        if (currentIndex < 7) txtChecks[currentIndex + 1].Focus();
                        else txtOtherText.Focus();
                    }
                };

                pnlInput.Controls.Add(lbl);
                pnlInput.Controls.Add(txtChecks[i]);
            }
            yPos += 4 * 40 + 20;

            pnlInput.Controls.Add(new Label { Text = "3. ⑨기타 (입력 후 Enter = 저장&다음사진):", AutoSize = true, Location = new Point(20, yPos) }); yPos += 25;
            pnlInput.Controls.Add(txtOtherText); txtOtherText.Location = new Point(20, yPos); yPos += 50;
            txtOtherText.KeyDown += TxtOtherText_KeyDown;

            pnlInput.Controls.Add(lblStatus); lblStatus.Location = new Point(20, yPos); yPos += 60;
            pnlInput.Controls.Add(btnExportExcel); btnExportExcel.Location = new Point(20, yPos);
            btnExportExcel.Click += BtnExportExcel_Click;

            // 우측 그리드
            dgvData = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, SelectionMode = DataGridViewSelectionMode.FullRowSelect, RowHeadersVisible = false };
            dgvData.Columns.Add("Id", "학번");
            dgvData.Columns.Add("Checks", "체크된 번호");
            dgvData.Columns.Add("Text", "기타 의견");

            mainLayout.Controls.Add(imageLayout, 0, 0);
            mainLayout.Controls.Add(pnlInput, 1, 0);
            mainLayout.Controls.Add(dgvData, 2, 0);

            this.Controls.Add(mainLayout);
        }

        private void LoadPdfOnStartup()
        {
            MessageBox.Show("설문지가 스캔된 200장짜리 PDF 파일을 선택해주세요.", "PDF 열기");
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "PDF Files|*.pdf" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    pdfDocument = PdfDocument.Load(ofd.FileName);

                    // PDF의 총 페이지 수만큼 대기열(Queue) 리스트 생성 (예: 0페이지 ~ 199페이지)
                    pagesToProcess = Enumerable.Range(0, pdfDocument.PageCount).ToList();

                    ShowCurrentPages();
                }
            }
        }

        // PDF 페이지를 선명한 화질(150 DPI)의 이미지로 변환하여 띄워주는 핵심 로직
        private void ShowCurrentPages()
        {
            if (picLeft.Image != null) { picLeft.Image.Dispose(); picLeft.Image = null; }
            if (picRight.Image != null) { picRight.Image.Dispose(); picRight.Image = null; }

            if (pagesToProcess.Count == 0)
            {
                lblImageProgress.Text = "모든 PDF 페이지 입력 완료!";
                MessageBox.Show("대기 중인 모든 페이지의 처리가 끝났습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 첫 번째(왼쪽) 장 렌더링
            int leftPageIndex = pagesToProcess[0];
            picLeft.Image = RenderPdfPage(leftPageIndex);

            // 두 번째(오른쪽) 장 렌더링 (대기열에 2장 이상 남아있을 경우에만)
            if (pagesToProcess.Count > 1)
            {
                int rightPageIndex = pagesToProcess[1];
                picRight.Image = RenderPdfPage(rightPageIndex);
                lblImageProgress.Text = $"남은 페이지: {pagesToProcess.Count} 장";
            }
            else
            {
                lblImageProgress.Text = $"남은 페이지: 1 장 (마지막)";
            }
        }

        private Image RenderPdfPage(int pageIndex)
        {
            int dpi = 150; // 해상도 설정 (기본 72보다 선명하게)
            var size = pdfDocument.PageSizes[pageIndex];
            int width = (int)(size.Width * dpi / 72.0);
            int height = (int)(size.Height * dpi / 72.0);

            return pdfDocument.Render(pageIndex, width, height, dpi, dpi, false);
        }

        private void TxtStudentId_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; txtChecks[0].Focus(); }
        }

        private void TxtOtherText_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; SaveCurrentData(); }
        }

        private void SaveCurrentData()
        {
            string studentId = txtStudentId.Text.Trim();
            if (string.IsNullOrEmpty(studentId)) return;

            var data = new SurveyData { StudentId = studentId };
            List<int> checkedNumbers = new List<int>();

            for (int i = 0; i < 8; i++)
            {
                string input = txtChecks[i].Text.Trim();
                if (input == "1" || input.ToLower() == "o" || input.ToLower() == "v")
                {
                    data.CheckedItems[i] = true;
                    checkedNumbers.Add(i + 1);
                }
            }

            data.OtherText = txtOtherText.Text.Trim();
            surveyDict[studentId] = data;
            string checkStr = checkedNumbers.Count > 0 ? string.Join(", ", checkedNumbers) : "없음";

            using (StreamWriter sw = new StreamWriter(csvFilePath, true, Encoding.UTF8))
            {
                sw.WriteLine($"{data.StudentId},\"{checkStr}\",\"{data.OtherText}\"");
            }

            dgvData.Rows.Insert(0, studentId, checkStr, data.OtherText);
            dgvData.Rows[0].Selected = true;
            lblStatus.Text = $"[{studentId}] 저장 완료! (현재 {surveyDict.Count}명)";

            txtStudentId.Clear();
            for (int i = 0; i < 8; i++) txtChecks[i].Clear();
            txtOtherText.Clear();
            txtStudentId.Focus();

            // ★ 정상 입력 완료 시, 앞의 두 장을 리스트에서 빼버리고 다음 사진 로드
            if (pagesToProcess.Count > 0) pagesToProcess.RemoveAt(0); // 왼쪽 제거
            if (pagesToProcess.Count > 0) pagesToProcess.RemoveAt(0); // 오른쪽 제거 (원래 1번이었던 것이 0번이 되므로 또 0번 제거)

            ShowCurrentPages();
        }

        private void BtnExportExcel_Click(object sender, EventArgs e)
        {
            if (surveyDict.Count == 0) { MessageBox.Show("입력된 데이터가 없습니다."); return; }
            OpenFileDialog openFileDialog = new OpenFileDialog { Title = "원본 엑셀 템플릿 선택", Filter = "Excel Files|*.xlsx" };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string templatePath = openFileDialog.FileName;
                string outputPath = Path.Combine(Path.GetDirectoryName(templatePath), "수요조사결과_완료.xlsx");

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
                                worksheet.Cell(row, 19).Value = data.OtherText;
                            }
                        }
                        workbook.SaveAs(outputPath);
                    }
                    MessageBox.Show($"변환 완료!\n위치: {outputPath}", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"오류: 엑셀 파일이 열려있다면 닫고 다시 시도하세요.\n\n{ex.Message}", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }

    public class SurveyData
    {
        public string StudentId { get; set; }
        public bool[] CheckedItems { get; set; } = new bool[8];
        public string OtherText { get; set; }
    }
}