using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Windows.Forms;

namespace ProcessCapability
{
    // Process Capability Report (공정능력분석) 메인 화면
    // 왼쪽: 입력 패널 / 오른쪽: 분석 리포트
    public class MainForm : Form
    {
        const int InputWidth = 320;
        const string DefaultData = "48.5 48.1 48.9 48.6 48.1 47.9 48.6 47.6 47.9 48.9 48.3 48.1 47.5 48.3 48.6 48.2 48.6 48.7 48.2 47.6";

        TextBox txtTitle, txtLsl, txtUsl, txtBinWidth;
        ComboBox cboChartType;
        FlowLayoutPanel groupsContainer;
        FlowLayoutPanel saveToolbar;
        ReportView reportView;
        int groupCounter = 1;

        public MainForm()
        {
            Text = "Process Capability Report (공정능력분석)";
            Font = new Font("Malgun Gothic", 9f);
            BackColor = ColorTranslator.FromHtml("#f5f6f8");
            ClientSize = new Size(1320, 820);
            MinimumSize = new Size(1150, 640);
            StartPosition = FormStartPosition.CenterScreen;

            BuildLayout();
        }

        // ======================================================================
        //  화면 구성
        // ======================================================================
        void BuildLayout()
        {
            // ----- 왼쪽 입력 패널 -----
            var input = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                Width = InputWidth + 40,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(16, 14, 16, 14),
                BackColor = Color.White
            };

            var header = new Label
            {
                Text = "Process Capability Report",
                Font = new Font("Malgun Gothic", 13f, FontStyle.Bold),
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 14)
            };
            input.Controls.Add(header);

            txtTitle = AddTextField(input, "리포트 제목", "");
            SetCue(txtTitle, "예: point 1의 공정능력");

            txtLsl = AddTextField(input, "하한 규격 (LSL)", "45");
            txtUsl = AddTextField(input, "상한 규격 (USL)", "50");
            input.Controls.Add(Hint("※ LSL 또는 USL 중 한쪽 칸을 비워두면 편측 규격(단측 검정)으로 계산됩니다."));
            txtBinWidth = AddTextField(input, "막대(Bin) 너비 (비워두면 자동)", "0.1");

            input.Controls.Add(FieldLabel("차트 종류"));
            cboChartType = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = InputWidth, Margin = new Padding(0, 0, 0, 12) };
            cboChartType.Items.AddRange(new object[] { "히스토그램 (분포)", "박스플롯 (Box Plot)" });
            cboChartType.SelectedIndex = 0;
            input.Controls.Add(cboChartType);

            input.Controls.Add(FieldLabel("측정 데이터 (그룹별로 입력, 박스플롯 비교에 사용됩니다)"));
            groupsContainer = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(InputWidth, 0),
                Margin = new Padding(0)
            };
            groupsContainer.Controls.Add(new GroupRow("Group 1", DefaultData));
            input.Controls.Add(groupsContainer);

            var groupButtons = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 0, 0, 4) };
            var btnAdd = StyledButton("+ 그룹 추가", "#eef2ff", "#c7d2fe", "#3730a3");
            btnAdd.Click += (s, e) => AddGroup();
            var btnReset = StyledButton("그룹 초기화", "#fff7ed", "#fdba74", "#9a3412");
            btnReset.Click += (s, e) => ResetGroups();
            groupButtons.Controls.Add(btnAdd);
            groupButtons.Controls.Add(btnReset);
            input.Controls.Add(groupButtons);

            input.Controls.Add(Hint("※ 줄바꿈(엔터), 탭, 띄어쓰기, 쉼표 모두 자동으로 인식합니다. 그룹을 여러 개 추가하면 박스플롯에서 함께 비교됩니다."));

            var btnAnalyze = new Button
            {
                Text = "분석 실행 (Analyze)",
                Width = InputWidth,
                Height = 44,
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorTranslator.FromHtml("#1d4ed8"),
                ForeColor = Color.White,
                Font = new Font("Malgun Gothic", 11f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 6, 0, 0)
            };
            btnAnalyze.FlatAppearance.BorderSize = 0;
            btnAnalyze.Click += (s, e) => AnalyzeData();
            input.Controls.Add(btnAnalyze);

            // ----- 오른쪽 리포트 영역 -----
            var reportArea = new Panel { Dock = DockStyle.Fill, Padding = new Padding(16) };

            saveToolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 40,
                FlowDirection = FlowDirection.RightToLeft,
                Visible = false
            };
            var btnSaveReport = StyledButton("전체 리포트 저장 (PNG)", "#ffffff", "#c9cdd6", "#1f2430");
            btnSaveReport.Click += (s, e) => SaveFullReport();
            var btnSaveChart = StyledButton("그래프만 저장 (PNG)", "#ffffff", "#c9cdd6", "#1f2430");
            btnSaveChart.Click += (s, e) => SaveChartOnly();
            saveToolbar.Controls.Add(btnSaveReport);
            saveToolbar.Controls.Add(btnSaveChart);

            reportView = new ReportView { Dock = DockStyle.Fill };

            reportArea.Controls.Add(reportView);
            reportArea.Controls.Add(saveToolbar);

            Controls.Add(reportArea);
            Controls.Add(input);
        }

        static Label FieldLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                MaximumSize = new Size(InputWidth, 0),
                Font = new Font("Malgun Gothic", 9f, FontStyle.Bold),
                ForeColor = ColorTranslator.FromHtml("#374151"),
                Margin = new Padding(0, 0, 0, 4)
            };
        }

        static Label Hint(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                MaximumSize = new Size(InputWidth, 0),
                ForeColor = ColorTranslator.FromHtml("#16a34a"),
                Font = new Font("Malgun Gothic", 8.5f),
                Margin = new Padding(0, 0, 0, 12)
            };
        }

        static TextBox AddTextField(Control parent, string label, string value)
        {
            parent.Controls.Add(FieldLabel(label));
            var tb = new TextBox { Text = value, Width = InputWidth, Margin = new Padding(0, 0, 0, 12) };
            parent.Controls.Add(tb);
            return tb;
        }

        static Button StyledButton(string text, string back, string border, string fore)
        {
            var b = new Button
            {
                Text = text,
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorTranslator.FromHtml(back),
                ForeColor = ColorTranslator.FromHtml(fore),
                Font = new Font("Malgun Gothic", 9f, FontStyle.Bold),
                Padding = new Padding(6, 2, 6, 2),
                Cursor = Cursors.Hand
            };
            b.FlatAppearance.BorderColor = ColorTranslator.FromHtml(border);
            return b;
        }

        // TextBox placeholder (EM_SETCUEBANNER)
        static void SetCue(TextBox tb, string cue)
        {
            tb.HandleCreated += (s, e) => SendMessage(tb.Handle, 0x1501, (IntPtr)1, cue);
        }

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

        // ======================================================================
        //  그룹 관리
        // ======================================================================
        void AddGroup()
        {
            groupCounter++;
            groupsContainer.Controls.Add(new GroupRow("Group " + groupCounter, ""));
        }

        void ResetGroups()
        {
            if (MessageBox.Show("모든 그룹 데이터를 지우고 그룹 1개만 남기시겠습니까?", "그룹 초기화",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;

            groupCounter = 1;
            foreach (Control c in groupsContainer.Controls.Cast<Control>().ToList())
            {
                groupsContainer.Controls.Remove(c);
                c.Dispose();
            }
            groupsContainer.Controls.Add(new GroupRow("Group 1", ""));
        }

        List<DataGroup> GetGroups()
        {
            var groups = new List<DataGroup>();
            int idx = 0;
            foreach (GroupRow row in groupsContainer.Controls.OfType<GroupRow>())
            {
                idx++;
                double[] data = CapabilityCalculator.ParseData(row.DataText);
                if (data.Length > 0)
                    groups.Add(new DataGroup(row.GroupName != "" ? row.GroupName : "Group " + idx, data));
            }
            return groups;
        }

        // ======================================================================
        //  분석
        // ======================================================================
        void AnalyzeData()
        {
            double? lsl, usl;
            if (!CapabilityCalculator.TryParseSpec(txtLsl.Text, out lsl)) { Alert("LSL 값이 올바른 숫자가 아닙니다."); return; }
            if (!CapabilityCalculator.TryParseSpec(txtUsl.Text, out usl)) { Alert("USL 값이 올바른 숫자가 아닙니다."); return; }

            List<DataGroup> groups = GetGroups();

            if (!lsl.HasValue && !usl.HasValue) { Alert("LSL 또는 USL 중 최소 하나는 입력해야 합니다."); return; }
            if (groups.Count == 0) { Alert("데이터가 없습니다. 최소 1개 그룹에 데이터를 입력해주세요."); return; }

            double[] data = groups.SelectMany(g => g.Data).ToArray();
            if (data.Length < 2) { Alert("데이터가 부족합니다."); return; }

            CapabilityResult result = CapabilityCalculator.Calculate(groups, lsl, usl);

            double? binWidth;
            CapabilityCalculator.TryParseSpec(txtBinWidth.Text, out binWidth); // 잘못된 값이면 자동

            string title = txtTitle.Text.Trim();
            if (cboChartType.SelectedIndex == 1)
                reportView.ShowBoxPlot(result, groups, title);
            else
                reportView.ShowHistogram(result, data, binWidth, title);

            saveToolbar.Visible = true;
        }

        void Alert(string message)
        {
            MessageBox.Show(this, message, "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        // ======================================================================
        //  PNG 저장
        // ======================================================================
        void SaveChartOnly()
        {
            if (!reportView.HasResult) { Alert("먼저 분석을 실행해주세요."); return; }
            using (Bitmap bmp = reportView.ExportChart())
                SavePng(bmp, "capability_chart_" + TimestampSuffix() + ".png");
        }

        void SaveFullReport()
        {
            if (!reportView.HasResult) { Alert("먼저 분석을 실행해주세요."); return; }
            try
            {
                using (Bitmap bmp = reportView.ExportReport())
                    SavePng(bmp, "capability_report_" + TimestampSuffix() + ".png");
            }
            catch (Exception ex)
            {
                Alert("리포트 저장 중 오류가 발생했습니다.\n" + ex.Message);
            }
        }

        void SavePng(Bitmap bmp, string defaultName)
        {
            using (var dlg = new SaveFileDialog
            {
                FileName = defaultName,
                Filter = "PNG 이미지 (*.png)|*.png",
                DefaultExt = "png"
            })
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                    bmp.Save(dlg.FileName, ImageFormat.Png);
            }
        }

        static string TimestampSuffix()
        {
            return DateTime.Now.ToString("yyyyMMdd_HHmmss");
        }
    }
}
