using System;
using System.Drawing;
using System.Windows.Forms;

namespace ProcessCapability
{
    // 그룹 1개 입력 행 (그룹명 + 삭제 버튼 + 측정 데이터)
    public class GroupRow : UserControl
    {
        readonly TextBox txtName;
        readonly TextBox txtData;
        readonly Button btnRemove;

        public GroupRow(string name, string data)
        {
            BackColor = ColorTranslator.FromHtml("#fbfbfc");
            BorderStyle = BorderStyle.FixedSingle;
            Padding = new Padding(8);
            Margin = new Padding(0, 0, 0, 8);
            Size = new Size(320, 118);

            txtName = new TextBox
            {
                Text = name,
                Font = new Font("Malgun Gothic", 9f, FontStyle.Bold),
                Location = new Point(8, 8),
                Width = 236,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            btnRemove = new Button
            {
                Text = "삭제",
                Location = new Point(250, 7),
                Size = new Size(58, 25),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorTranslator.FromHtml("#fff1f1"),
                ForeColor = ColorTranslator.FromHtml("#c0392b"),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Cursor = Cursors.Hand
            };
            btnRemove.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#f3c2c2");
            btnRemove.Click += (s, e) =>
            {
                var parent = Parent;
                if (parent != null) parent.Controls.Remove(this);
                Dispose();
            };

            txtData = new TextBox
            {
                Text = data,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Location = new Point(8, 40),
                Size = new Size(300, 68),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };

            Controls.Add(txtName);
            Controls.Add(btnRemove);
            Controls.Add(txtData);
        }

        public string GroupName { get { return txtName.Text.Trim(); } }
        public string DataText { get { return txtData.Text; } }
    }
}
