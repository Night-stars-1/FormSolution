using System.Windows.Forms;

public class ProgressForm : Form
{
    private ProgressBar progressBar;

    public ProgressForm()
    {
        this.Text = "请稍候";
        this.Width = 300;
        this.Height = 150;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.ControlBox = false; // 禁用关闭按钮
        this.TopMost = true; // 窗口始终在最前

        Label label = new Label
        {
            Text = "正在处理，请稍候...",
            Dock = DockStyle.Top,
            Height = 30,
            TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        };

        progressBar = new ProgressBar
        {
            Dock = DockStyle.Bottom,
            Style = ProgressBarStyle.Marquee,
            Height = 20,
            MarqueeAnimationSpeed = 30
        };

        this.Controls.Add(progressBar);
        this.Controls.Add(label);
    }
}
