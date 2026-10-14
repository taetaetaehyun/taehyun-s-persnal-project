using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace ProcessCapability
{
    public enum ChartKind
    {
        Histogram,
        BoxPlot
    }

    // Process Capability Report 대시보드 (Process Data / 차트 / Capability 박스)
    // 모든 요소를 GDI+로 직접 그려서 화면 표시와 PNG 저장 결과가 동일하게 나오도록 함
    public class ReportView : Control
    {
        // ===== 색상 / 폰트 =====
        static readonly Color DashboardBack = ColorTranslator.FromHtml("#e8e8e8");
        static readonly Color DashboardBorder = ColorTranslator.FromHtml("#cfd2d8");
        static readonly Color GridColor = ColorTranslator.FromHtml("#dcdcdc");
        static readonly Color SpecColor = ColorTranslator.FromHtml("#c00000");
        static readonly Color BarColor = ColorTranslator.FromHtml("#bccde5");
        static readonly Color CurveColor = ColorTranslator.FromHtml("#800000");
        static readonly Color TitleText = ColorTranslator.FromHtml("#1f2430");
        static readonly Color LabelText = ColorTranslator.FromHtml("#333333");
        static readonly Color BoxTitleLine = ColorTranslator.FromHtml("#cccccc");
        static readonly Color[] Palette =
        {
            ColorTranslator.FromHtml("#4472c4"), ColorTranslator.FromHtml("#ed7d31"),
            ColorTranslator.FromHtml("#a5a5a5"), ColorTranslator.FromHtml("#ffc000"),
            ColorTranslator.FromHtml("#70ad47"), ColorTranslator.FromHtml("#5b9bd5"),
            ColorTranslator.FromHtml("#c00000"), ColorTranslator.FromHtml("#7030a0")
        };

        const string LatinFont = "Tahoma";
        const string HangulFont = "Malgun Gothic";

        // ===== 레이아웃 상수 (px) =====
        const int Pad = 14;
        const int SideColWidth = 170;
        const int TitleHeight = 30;
        const int MinChartHeight = 380;
        const int BoxTitleHeight = 26;
        const int BoxRowHeight = 19;

        // ===== 분석 데이터 =====
        CapabilityResult result;
        ChartKind chartKind;
        string reportTitle = "";

        // 히스토그램
        double histMin, histMax, histBinWidth;
        int[] frequencies;
        PointF[] curve; // (x값, y값) - 픽셀 아님

        // 박스플롯
        List<BoxStat> boxStats;

        class BoxStat
        {
            public string Name;
            public Color Color;
            public double CenterX, Q1, Q3, Median, WhiskerLow, WhiskerHigh;
            public PointF[] Points; // 지터 적용된 (x, 값)
        }

        public ReportView()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
            BackColor = Color.White;
        }

        public bool HasResult { get { return result != null; } }

        // ======================================================================
        //  데이터 설정
        // ======================================================================
        public void ShowHistogram(CapabilityResult r, double[] data, double? userBinWidth, string title)
        {
            result = r;
            reportTitle = title ?? "";
            chartKind = ChartKind.Histogram;

            var specs = SpecValues();
            double std = r.StdWithin;
            double pad = std > 0 ? std * 1.5 : 0;
            histMin = data.Concat(specs).Min() - pad;
            histMax = data.Concat(specs).Max() + pad;
            if (histMax - histMin <= 0) { histMin -= 1; histMax += 1; }

            double binWidth = userBinWidth ?? 0;
            int binCount;
            if (binWidth <= 0 || double.IsNaN(binWidth) || (histMax - histMin) / binWidth > 2000)
            {
                binCount = Math.Max(7, (int)Math.Ceiling(Math.Sqrt(data.Length)));
                binWidth = (histMax - histMin) / binCount;
            }
            else
            {
                binCount = (int)Math.Ceiling((histMax - histMin) / binWidth);
            }
            histBinWidth = binWidth;

            frequencies = new int[binCount];
            for (int i = 0; i < binCount; i++)
            {
                double start = histMin + i * binWidth;
                double end = start + binWidth;
                frequencies[i] = data.Count(v => v >= start && v < end);
            }

            // 정규분포 곡선 (Within 표준편차 기준)
            if (std > 0)
            {
                const int curvePoints = 120;
                double step = (histMax - histMin) / (curvePoints - 1);
                double scale = data.Length * binWidth;
                curve = new PointF[curvePoints];
                for (int i = 0; i < curvePoints; i++)
                {
                    double x = histMin + i * step;
                    curve[i] = new PointF((float)x, (float)(scale * CapabilityCalculator.NormalPdf(x, r.Mean, std)));
                }
            }
            else
            {
                curve = null;
            }

            Invalidate();
        }

        public void ShowBoxPlot(CapabilityResult r, IList<DataGroup> groups, string title)
        {
            result = r;
            reportTitle = title ?? "";
            chartKind = ChartKind.BoxPlot;

            var rnd = new Random();
            const double jitterHalf = 0.16;
            boxStats = new List<BoxStat>();

            for (int idx = 0; idx < groups.Count; idx++)
            {
                double[] sorted = groups[idx].Data.OrderBy(v => v).ToArray();
                double median = CapabilityCalculator.Percentile(sorted, 50);
                double q1 = CapabilityCalculator.Percentile(sorted, 25);
                double q3 = CapabilityCalculator.Percentile(sorted, 75);
                double iqr = q3 - q1;
                double lowerFence = q1 - 1.5 * iqr;
                double upperFence = q3 + 1.5 * iqr;
                var withinLow = sorted.Where(v => v >= lowerFence).ToArray();
                var withinHigh = sorted.Where(v => v <= upperFence).ToArray();
                double centerX = idx + 0.5;

                boxStats.Add(new BoxStat
                {
                    Name = groups[idx].Name,
                    Color = Palette[idx % Palette.Length],
                    CenterX = centerX,
                    Q1 = q1,
                    Q3 = q3,
                    Median = median,
                    WhiskerLow = withinLow.Length > 0 ? withinLow.Min() : sorted[0],
                    WhiskerHigh = withinHigh.Length > 0 ? withinHigh.Max() : sorted[sorted.Length - 1],
                    // 지터는 분석 시 한 번만 계산 (다시 그려도 점 위치가 바뀌지 않도록)
                    Points = groups[idx].Data
                        .Select(v => new PointF((float)(centerX + (rnd.NextDouble() - 0.5) * jitterHalf * 2), (float)v))
                        .ToArray()
                });
            }

            Invalidate();
        }

        IEnumerable<double> SpecValues()
        {
            var specs = new List<double>();
            if (result.Lsl.HasValue) specs.Add(result.Lsl.Value);
            if (result.Usl.HasValue) specs.Add(result.Usl.Value);
            return specs;
        }

        // ======================================================================
        //  PNG 내보내기
        // ======================================================================

        // 전체 리포트 (2배 해상도)
        public Bitmap ExportReport(float scale = 2f)
        {
            Size size = ClientSize;
            var bmp = new Bitmap((int)(size.Width * scale), (int)(size.Height * scale), PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.ScaleTransform(scale, scale);
                RenderDashboard(g, size);
            }
            return bmp;
        }

        // 그래프만 (흰 배경, 2배 해상도)
        public Bitmap ExportChart(float scale = 2f)
        {
            Rectangle wrap = GetLayout(ClientSize).Chart;
            Rectangle content = ChartContentRect(wrap);
            var bmp = new Bitmap((int)(content.Width * scale), (int)(content.Height * scale), PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                g.ScaleTransform(scale, scale);
                g.TranslateTransform(-content.X, -content.Y);
                SetQuality(g);
                RenderChart(g, content);
            }
            return bmp;
        }

        // ======================================================================
        //  그리기
        // ======================================================================
        protected override void OnPaint(PaintEventArgs e)
        {
            if (result == null)
            {
                e.Graphics.Clear(BackColor);
                using (var f = new Font(HangulFont, 14, GraphicsUnit.Pixel))
                {
                    TextRenderer.DrawText(e.Graphics, "데이터를 입력하고 [분석 실행 (Analyze)] 버튼을 눌러주세요.", f,
                        ClientRectangle, Color.Gray, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                }
                return;
            }
            RenderDashboard(e.Graphics, ClientSize);
        }

        struct DashLayout
        {
            public Rectangle Title, Left, Chart, RightTop, RightBottom;
        }

        DashLayout GetLayout(Size size)
        {
            var l = new DashLayout();
            int y = Pad;
            if (!string.IsNullOrEmpty(reportTitle))
            {
                l.Title = new Rectangle(Pad, y, size.Width - Pad * 2, TitleHeight);
                y += TitleHeight + 4;
            }

            int chartHeight = Math.Max(MinChartHeight, size.Height - y - Pad);
            int chartX = Pad + SideColWidth + Pad;
            int chartWidth = Math.Max(200, size.Width - (SideColWidth + Pad) * 2 - Pad * 2);
            int rightX = chartX + chartWidth + Pad;

            l.Left = new Rectangle(Pad, y, SideColWidth, BoxHeight(6));
            l.Chart = new Rectangle(chartX, y, chartWidth, chartHeight);
            l.RightTop = new Rectangle(rightX, y, SideColWidth, BoxHeight(5));
            l.RightBottom = new Rectangle(rightX, l.RightTop.Bottom + 10, SideColWidth, BoxHeight(5));
            return l;
        }

        static int BoxHeight(int rows)
        {
            return 10 + BoxTitleHeight + rows * BoxRowHeight + 6;
        }

        void RenderDashboard(Graphics g, Size size)
        {
            SetQuality(g);
            g.Clear(DashboardBack);
            using (var pen = new Pen(DashboardBorder))
                g.DrawRectangle(pen, 0, 0, size.Width - 1, size.Height - 1);

            DashLayout l = GetLayout(size);
            var r = result;

            if (!string.IsNullOrEmpty(reportTitle))
            {
                using (var f = new Font(FontFor(reportTitle), 18, FontStyle.Bold, GraphicsUnit.Pixel))
                using (var b = new SolidBrush(TitleText))
                using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                    g.DrawString(reportTitle, f, b, l.Title, sf);
            }

            DrawDataBox(g, l.Left, "Process Data", new[]
            {
                Row("LSL", r.Lsl.HasValue ? FormatSpec(r.Lsl.Value) : "*"),
                Row("USL", r.Usl.HasValue ? FormatSpec(r.Usl.Value) : "*"),
                Row("Sample Mean", Fmt(r.Mean, 4)),
                Row("Sample N", r.GroupCount > 1 ? string.Format("{0} ({1}개 그룹 합산)", r.N, r.GroupCount) : r.N.ToString()),
                Row("StDev (Within)", Fmt(r.StdWithin, 5)),
                Row("StDev (Overall)", Fmt(r.StdOverall, 5))
            });

            DrawDataBox(g, l.RightTop, "Potential (Within) Capability", new[]
            {
                Row("Cp", Fmt(r.Cp, 3)),
                Row("CPL", Fmt(r.Cpl, 3)),
                Row("CPU", Fmt(r.Cpu, 3)),
                Row("Cpk", Fmt(r.Cpk, 3)),
                Row("CCpk", Fmt(r.CCpk, 3))
            });

            DrawDataBox(g, l.RightBottom, "Overall Capability", new[]
            {
                Row("Pp", Fmt(r.Pp, 3)),
                Row("PPL", Fmt(r.Ppl, 3)),
                Row("PPU", Fmt(r.Ppu, 3)),
                Row("Ppk", Fmt(r.Ppk, 3)),
                Row("Cpm", Fmt(r.Cpm, 3))
            });

            // 차트 영역 (흰 배경 + 검은 테두리)
            g.FillRectangle(Brushes.White, l.Chart);
            g.DrawRectangle(Pens.Black, l.Chart.X, l.Chart.Y, l.Chart.Width - 1, l.Chart.Height - 1);
            RenderChart(g, ChartContentRect(l.Chart));
        }

        static Rectangle ChartContentRect(Rectangle wrap)
        {
            return Rectangle.FromLTRB(wrap.Left + 16, wrap.Top + 14, wrap.Right - 16, wrap.Bottom - 6);
        }

        static KeyValuePair<string, string> Row(string label, string value)
        {
            return new KeyValuePair<string, string>(label, value);
        }

        void DrawDataBox(Graphics g, Rectangle rect, string title, KeyValuePair<string, string>[] rows)
        {
            g.FillRectangle(Brushes.White, rect);
            g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width - 1, rect.Height - 1);

            int x = rect.X + 12, right = rect.Right - 12;
            int y = rect.Y + 10;

            using (var tf = new Font(LatinFont, 12, FontStyle.Bold, GraphicsUnit.Pixel))
            using (var center = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Near })
                g.DrawString(title, tf, Brushes.Black, new RectangleF(rect.X, y, rect.Width, 18), center);

            using (var linePen = new Pen(BoxTitleLine))
                g.DrawLine(linePen, x, y + BoxTitleHeight - 8, right, y + BoxTitleHeight - 8);
            y += BoxTitleHeight;

            using (var labelBrush = new SolidBrush(LabelText))
            using (var near = new StringFormat { Alignment = StringAlignment.Near, FormatFlags = StringFormatFlags.NoWrap })
            using (var far = new StringFormat { Alignment = StringAlignment.Far, FormatFlags = StringFormatFlags.NoWrap })
            {
                foreach (var row in rows)
                {
                    using (var lf = new Font(LatinFont, 12, GraphicsUnit.Pixel))
                    using (var vf = new Font(FontFor(row.Value), 12, FontStyle.Bold, GraphicsUnit.Pixel))
                    {
                        g.DrawString(row.Key, lf, labelBrush, x, y, near);
                        g.DrawString(row.Value, vf, Brushes.Black, new RectangleF(x, y, right - x, BoxRowHeight), far);
                    }
                    y += BoxRowHeight;
                }
            }
        }

        void RenderChart(Graphics g, Rectangle area)
        {
            GraphicsState state = g.Save();
            g.SetClip(area);
            if (chartKind == ChartKind.BoxPlot)
                RenderBoxPlot(g, area);
            else
                RenderHistogram(g, area);
            g.Restore(state);
        }

        // ---------------------------------------------------------------- 히스토그램
        void RenderHistogram(Graphics g, Rectangle area)
        {
            using (var tickFont = new Font(LatinFont, 10, GraphicsUnit.Pixel))
            using (var titleFont = new Font(HangulFont, 12, GraphicsUnit.Pixel))
            using (var legendFont = new Font(HangulFont, 11, GraphicsUnit.Pixel))
            using (var specFont = new Font(LatinFont, 12, FontStyle.Bold, GraphicsUnit.Pixel))
            {
                var legend = new List<LegendItem>
                {
                    new LegendItem("데이터 빈도 (Histogram)", BarColor, Color.Black, 1),
                    new LegendItem("정규분포 (Normal Curve)", Color.FromArgb(40, CurveColor), CurveColor, 2)
                };
                int legendBottom = DrawLegend(g, legend, legendFont, area);

                // y축 범위 (정수 눈금)
                double yDataMax = frequencies.Length > 0 ? frequencies.Max() : 1;
                if (curve != null) yDataMax = Math.Max(yDataMax, curve.Max(p => p.Y));
                if (yDataMax <= 0) yDataMax = 1;
                double yStep = Math.Max(1, Math.Ceiling(NiceStep(yDataMax, 6)));
                double yMax = Math.Ceiling(yDataMax / yStep) * yStep;
                var yTicks = new List<double>();
                for (double v = 0; v <= yMax + yStep * 1e-9; v += yStep) yTicks.Add(v);

                var xTicks = NiceTicks(histMin, histMax, Math.Max(4, area.Width / 70));
                int xDecimals = Decimals(xTicks.Count > 1 ? xTicks[1] - xTicks[0] : 1);

                float yTickWidth = yTicks.Max(v => g.MeasureString(FormatTick(v, 0), tickFont).Width);
                float plotLeft = area.Left + 18 + yTickWidth + 6;
                float plotRight = area.Right - 4;
                float plotTop = legendBottom + 24;  // LSL/USL 라벨 자리
                float plotBottom = area.Bottom - 40; // x축 눈금 + 제목
                var plot = RectangleF.FromLTRB(plotLeft, plotTop, plotRight, plotBottom);

                Func<double, float> px = v => (float)(plot.Left + (v - histMin) / (histMax - histMin) * plot.Width);
                Func<double, float> py = v => (float)(plot.Bottom - v / yMax * plot.Height);

                // 그리드 + 눈금
                using (var gridPen = new Pen(GridColor))
                using (var center = new StringFormat { Alignment = StringAlignment.Center })
                using (var right = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center })
                {
                    foreach (double v in xTicks)
                    {
                        float x = px(v);
                        g.DrawLine(gridPen, x, plot.Top, x, plot.Bottom);
                        g.DrawString(FormatTick(v, xDecimals), tickFont, Brushes.Black, x, plot.Bottom + 4, center);
                    }
                    foreach (double v in yTicks)
                    {
                        float y = py(v);
                        g.DrawLine(gridPen, plot.Left, y, plot.Right, y);
                        g.DrawString(FormatTick(v, 0), tickFont, Brushes.Black, plot.Left - 4, y, right);
                    }
                    g.DrawString("측정값 (Value)", titleFont, Brushes.Black, (plot.Left + plot.Right) / 2, plot.Bottom + 20, center);
                }
                DrawRotatedText(g, "빈도수 (Frequency)", titleFont, area.Left + 8, (plot.Top + plot.Bottom) / 2);

                // 막대
                GraphicsState s = g.Save();
                g.SetClip(plot, CombineMode.Intersect);
                using (var barBrush = new SolidBrush(BarColor))
                {
                    for (int i = 0; i < frequencies.Length; i++)
                    {
                        if (frequencies[i] == 0) continue;
                        double start = histMin + i * histBinWidth;
                        float x1 = px(start), x2 = px(start + histBinWidth);
                        float y1 = py(frequencies[i]);
                        var bar = RectangleF.FromLTRB(x1, y1, x2, plot.Bottom);
                        g.FillRectangle(barBrush, bar);
                        g.DrawRectangle(Pens.Black, bar.X, bar.Y, bar.Width, bar.Height);
                    }
                }

                // 정규분포 곡선
                if (curve != null)
                {
                    using (var curvePen = new Pen(CurveColor, 2f) { LineJoin = LineJoin.Round })
                        g.DrawLines(curvePen, curve.Select(p => new PointF(px(p.X), py(p.Y))).ToArray());
                }
                g.Restore(s);

                // 축선
                g.DrawLine(Pens.Black, plot.Left, plot.Top, plot.Left, plot.Bottom);
                g.DrawLine(Pens.Black, plot.Left, plot.Bottom, plot.Right, plot.Bottom);

                // LSL / USL 세로 점선
                if (result.Lsl.HasValue) DrawVSpec(g, plot, px(result.Lsl.Value), "LSL", specFont);
                if (result.Usl.HasValue) DrawVSpec(g, plot, px(result.Usl.Value), "USL", specFont);
            }
        }

        void DrawVSpec(Graphics g, RectangleF plot, float x, string label, Font font)
        {
            if (x < plot.Left - 0.5f || x > plot.Right + 0.5f) return;
            using (var pen = new Pen(SpecColor, 1.5f) { DashPattern = new[] { 5f / 1.5f, 4f / 1.5f } })
                g.DrawLine(pen, x, plot.Top, x, plot.Bottom);
            using (var brush = new SolidBrush(SpecColor))
            using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Far })
                g.DrawString(label, font, brush, x, plot.Top - 4, sf);
        }

        // ---------------------------------------------------------------- 박스플롯
        void RenderBoxPlot(Graphics g, Rectangle area)
        {
            using (var tickFont = new Font(LatinFont, 10, GraphicsUnit.Pixel))
            using (var titleFont = new Font(HangulFont, 12, GraphicsUnit.Pixel))
            using (var legendFont = new Font(HangulFont, 11, GraphicsUnit.Pixel))
            using (var specFont = new Font(LatinFont, 11, FontStyle.Bold, GraphicsUnit.Pixel))
            {
                var legend = boxStats.Select(b => new LegendItem(b.Name, b.Color, b.Color, 1)).ToList();
                int legendBottom = DrawLegend(g, legend, legendFont, area);

                var all = boxStats.SelectMany(b => b.Points.Select(p => (double)p.Y)).Concat(SpecValues()).ToArray();
                double dataMin = all.Min(), dataMax = all.Max();
                double range = dataMax - dataMin;
                if (range == 0) range = 1;
                double yMin = dataMin - range * 0.1;
                double yMax = dataMax + range * 0.1;
                int nGroups = boxStats.Count;

                var yTicks = NiceTicks(yMin, yMax, Math.Max(4, area.Height / 45));
                int yDecimals = Decimals(yTicks.Count > 1 ? yTicks[1] - yTicks[0] : 1);

                float yTickWidth = yTicks.Count > 0 ? yTicks.Max(v => g.MeasureString(FormatTick(v, yDecimals), tickFont).Width) : 20;
                float plotLeft = area.Left + 10 + 18 + yTickWidth + 6;
                float plotRight = area.Right - 40;  // LSL/USL 라벨 자리
                float plotTop = legendBottom + 12;
                float plotBottom = area.Bottom - 10;
                var plot = RectangleF.FromLTRB(plotLeft, plotTop, plotRight, plotBottom);

                Func<double, float> px = v => (float)(plot.Left + v / nGroups * plot.Width);
                Func<double, float> py = v => (float)(plot.Bottom - (v - yMin) / (yMax - yMin) * plot.Height);

                using (var gridPen = new Pen(GridColor))
                using (var right = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center })
                {
                    foreach (double v in yTicks)
                    {
                        float y = py(v);
                        g.DrawLine(gridPen, plot.Left, y, plot.Right, y);
                        g.DrawString(FormatTick(v, yDecimals), tickFont, Brushes.Black, plot.Left - 4, y, right);
                    }
                }
                g.DrawLine(Pens.Black, plot.Left, plot.Top, plot.Left, plot.Bottom);
                DrawRotatedText(g, "측정값 (Value)", titleFont, area.Left + 18, (plot.Top + plot.Bottom) / 2);

                // 개별 데이터 점 (지터)
                foreach (var b in boxStats)
                {
                    using (var brush = new SolidBrush(b.Color))
                    {
                        foreach (var p in b.Points)
                            g.FillEllipse(brush, px(p.X) - 4, py(p.Y) - 4, 8, 8);
                    }
                }

                // 박스 + 수염
                foreach (var b in boxStats)
                {
                    float xc = px(b.CenterX);
                    float boxHalf = px(b.CenterX + 0.3) - xc;
                    float capHalf = boxHalf * 0.55f;
                    float yQ1 = py(b.Q1), yQ3 = py(b.Q3), yMed = py(b.Median);
                    float yWLow = py(b.WhiskerLow), yWHigh = py(b.WhiskerHigh);

                    using (var pen = new Pen(b.Color, 1.5f))
                    using (var fill = new SolidBrush(Color.FromArgb(64, b.Color)))
                    using (var medPen = new Pen(b.Color, 2f))
                    {
                        g.DrawLine(pen, xc, yQ3, xc, yWHigh);
                        g.DrawLine(pen, xc, yQ1, xc, yWLow);
                        g.DrawLine(pen, xc - capHalf, yWHigh, xc + capHalf, yWHigh);
                        g.DrawLine(pen, xc - capHalf, yWLow, xc + capHalf, yWLow);

                        var box = RectangleF.FromLTRB(xc - boxHalf, yQ3, xc + boxHalf, yQ1);
                        g.FillRectangle(fill, box);
                        g.DrawRectangle(pen, box.X, box.Y, box.Width, box.Height);
                        g.DrawLine(medPen, xc - boxHalf, yMed, xc + boxHalf, yMed);
                    }
                }

                // LSL / USL 가로 점선
                if (result.Lsl.HasValue) DrawHSpec(g, plot, py(result.Lsl.Value), "LSL", specFont);
                if (result.Usl.HasValue) DrawHSpec(g, plot, py(result.Usl.Value), "USL", specFont);
            }
        }

        void DrawHSpec(Graphics g, RectangleF plot, float y, string label, Font font)
        {
            if (y < plot.Top - 0.5f || y > plot.Bottom + 0.5f) return;
            using (var pen = new Pen(SpecColor, 1.5f) { DashPattern = new[] { 5f / 1.5f, 4f / 1.5f } })
                g.DrawLine(pen, plot.Left, y, plot.Right, y);
            using (var brush = new SolidBrush(SpecColor))
            using (var sf = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center })
                g.DrawString(label, font, brush, plot.Right + 4, y, sf);
        }

        // ---------------------------------------------------------------- 범례
        class LegendItem
        {
            public string Text;
            public Color Fill, Border;
            public float BorderWidth;

            public LegendItem(string text, Color fill, Color border, float borderWidth)
            {
                Text = text;
                Fill = fill;
                Border = border;
                BorderWidth = borderWidth;
            }
        }

        // 오른쪽 정렬 범례 (넘치면 줄바꿈). 범례 하단 y 좌표 반환
        static int DrawLegend(Graphics g, List<LegendItem> items, Font font, Rectangle area)
        {
            const int box = 14, gapText = 6, gapItem = 12, rowHeight = 18;
            var rows = new List<List<KeyValuePair<LegendItem, float>>> { new List<KeyValuePair<LegendItem, float>>() };
            float rowWidth = 0;

            foreach (var item in items)
            {
                float w = box + gapText + g.MeasureString(item.Text, font).Width;
                if (rows[rows.Count - 1].Count > 0 && rowWidth + gapItem + w > area.Width)
                {
                    rows.Add(new List<KeyValuePair<LegendItem, float>>());
                    rowWidth = 0;
                }
                if (rows[rows.Count - 1].Count > 0) rowWidth += gapItem;
                rows[rows.Count - 1].Add(new KeyValuePair<LegendItem, float>(item, w));
                rowWidth += w;
            }

            float y = area.Top;
            using (var sf = new StringFormat { LineAlignment = StringAlignment.Center })
            {
                foreach (var row in rows)
                {
                    float total = row.Sum(p => p.Value) + gapItem * Math.Max(0, row.Count - 1);
                    float x = area.Right - total;
                    foreach (var p in row)
                    {
                        var item = p.Key;
                        using (var fill = new SolidBrush(item.Fill))
                        using (var pen = new Pen(item.Border, item.BorderWidth))
                        {
                            g.FillRectangle(fill, x, y + (rowHeight - box) / 2f, box, box);
                            g.DrawRectangle(pen, x, y + (rowHeight - box) / 2f, box, box);
                        }
                        g.DrawString(item.Text, font, Brushes.Black, x + box + gapText, y + rowHeight / 2f, sf);
                        x += p.Value + gapItem;
                    }
                    y += rowHeight;
                }
            }
            return (int)Math.Ceiling(y);
        }

        // ======================================================================
        //  유틸
        // ======================================================================
        static void SetQuality(Graphics g)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            g.PixelOffsetMode = PixelOffsetMode.Half;
        }

        static void DrawRotatedText(Graphics g, string text, Font font, float x, float y)
        {
            GraphicsState s = g.Save();
            g.TranslateTransform(x, y);
            g.RotateTransform(-90);
            using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                g.DrawString(text, font, Brushes.Black, 0, 0, sf);
            g.Restore(s);
        }

        // 한글이 포함된 문자열은 맑은 고딕 사용 (Tahoma에는 한글 글리프가 없음)
        static string FontFor(string text)
        {
            foreach (char c in text)
                if (c >= 0xAC00 && c <= 0xD7A3 || c >= 0x3130 && c <= 0x318F) return HangulFont;
            return LatinFont;
        }

        static string Fmt(double? v, int decimals)
        {
            if (!v.HasValue || double.IsNaN(v.Value) || double.IsInfinity(v.Value)) return "*";
            return v.Value.ToString("F" + decimals, CultureInfo.InvariantCulture);
        }

        static string FormatSpec(double v)
        {
            return v.ToString("R", CultureInfo.InvariantCulture);
        }

        static string FormatTick(double v, int decimals)
        {
            if (Math.Abs(v) < 1e-12) v = 0;
            return v.ToString("F" + decimals, CultureInfo.InvariantCulture);
        }

        static double NiceStep(double range, int maxTicks)
        {
            double rough = range / Math.Max(1, maxTicks - 1);
            double mag = Math.Pow(10, Math.Floor(Math.Log10(rough)));
            double norm = rough / mag;
            double nice = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10;
            return nice * mag;
        }

        // [min, max] 범위 안의 보기 좋은 눈금
        static List<double> NiceTicks(double min, double max, int maxTicks)
        {
            var ticks = new List<double>();
            if (!(max > min)) return ticks;
            double step = NiceStep(max - min, maxTicks);
            double start = Math.Ceiling(min / step - 1e-9) * step;
            for (double v = start; v <= max + step * 1e-9; v += step)
                ticks.Add(Math.Round(v / step) * step);
            return ticks;
        }

        static int Decimals(double step)
        {
            for (int d = 0; d <= 8; d++)
            {
                double scaled = step * Math.Pow(10, d);
                if (Math.Abs(scaled - Math.Round(scaled)) < 1e-6 * Math.Max(1, scaled)) return d;
            }
            return 8;
        }
    }
}
