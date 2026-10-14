using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ProcessCapability
{
    // 측정 데이터 그룹 (박스플롯 비교용)
    public class DataGroup
    {
        public string Name;
        public double[] Data;

        public DataGroup(string name, double[] data)
        {
            Name = name;
            Data = data;
        }
    }

    // 공정능력 분석 결과 (규격이 없는 항목은 null → 화면에 '*' 표시)
    public class CapabilityResult
    {
        public double? Lsl;
        public double? Usl;
        public int N;
        public int GroupCount;
        public double Mean;
        public double StdWithin;
        public double StdOverall;

        public double? Cp, Cpl, Cpu, Cpk, CCpk;
        public double? Pp, Ppl, Ppu, Ppk, Cpm;
    }

    public static class CapabilityCalculator
    {
        // 이동범위(MR, span 2)용 관리도 상수
        const double D2 = 1.128;

        // 줄바꿈, 탭, 띄어쓰기, 쉼표 모두 구분자로 인식
        public static double[] ParseData(string raw)
        {
            var values = new List<double>();
            if (string.IsNullOrEmpty(raw)) return values.ToArray();

            foreach (string token in Regex.Split(raw, @"[\r\n\t, ]+"))
            {
                double v;
                if (double.TryParse(token.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out v))
                    values.Add(v);
            }
            return values.ToArray();
        }

        // 빈 문자열이면 null (편측 규격)
        public static bool TryParseSpec(string raw, out double? value)
        {
            value = null;
            raw = (raw ?? "").Trim();
            if (raw == "") return true;

            double v;
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out v)) return false;
            value = v;
            return true;
        }

        public static CapabilityResult Calculate(IList<DataGroup> groups, double? lsl, double? usl)
        {
            double[] data = groups.SelectMany(g => g.Data).ToArray();
            int n = data.Length;

            double mean = data.Average();
            double variance = data.Sum(v => Math.Pow(v - mean, 2)) / (n - 1);
            double sampleStd = Math.Sqrt(variance);
            double c4 = (4.0 * (n - 1)) / (4.0 * n - 3);
            double stdOverall = sampleStd / c4;

            double mrSum = 0;
            for (int i = 1; i < n; i++)
                mrSum += Math.Abs(data[i] - data[i - 1]);
            double mrAvg = mrSum / (n - 1);
            double stdWithin = mrAvg / D2;

            var r = new CapabilityResult
            {
                Lsl = lsl,
                Usl = usl,
                N = n,
                GroupCount = groups.Count,
                Mean = mean,
                StdWithin = stdWithin,
                StdOverall = stdOverall
            };

            // Potential (Within) Capability
            if (lsl.HasValue && usl.HasValue) r.Cp = (usl.Value - lsl.Value) / (6 * stdWithin);
            if (usl.HasValue) r.Cpu = (usl.Value - mean) / (3 * stdWithin);
            if (lsl.HasValue) r.Cpl = (mean - lsl.Value) / (3 * stdWithin);
            r.Cpk = MinOf(r.Cpu, r.Cpl);
            r.CCpk = r.Cp ?? r.Cpk;

            // Overall Capability
            if (lsl.HasValue && usl.HasValue) r.Pp = (usl.Value - lsl.Value) / (6 * stdOverall);
            if (usl.HasValue) r.Ppu = (usl.Value - mean) / (3 * stdOverall);
            if (lsl.HasValue) r.Ppl = (mean - lsl.Value) / (3 * stdOverall);
            r.Ppk = MinOf(r.Ppu, r.Ppl);
            r.Cpm = null; // 목표값(Target) 미지정

            return r;
        }

        static double? MinOf(double? a, double? b)
        {
            if (a.HasValue && b.HasValue) return Math.Min(a.Value, b.Value);
            return a ?? b;
        }

        // 선형 보간 백분위수 (JS 버전과 동일)
        public static double Percentile(double[] sorted, double p)
        {
            double idx = (p / 100.0) * (sorted.Length - 1);
            int lo = (int)Math.Floor(idx);
            int hi = (int)Math.Ceiling(idx);
            if (lo == hi) return sorted[lo];
            return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
        }

        public static double NormalPdf(double x, double mean, double std)
        {
            return (1 / (std * Math.Sqrt(2 * Math.PI))) * Math.Exp(-0.5 * Math.Pow((x - mean) / std, 2));
        }
    }
}
