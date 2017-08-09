using System;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace PowerPointAddIn1
{
    public partial class EvaluationChartForm : Form
    {
        public EvaluationChartForm(String question)
        {
            InitializeComponent();

            //====Bar Chart
            this.components = new System.ComponentModel.Container();
            ChartArea chartArea1 = new ChartArea();
            Legend legend2 = new Legend()
            { BackColor = Color.Green, ForeColor = Color.Black, Title = "Salary" };
            Chart pieChart = new Chart();
            Chart barChart = new Chart();

            chartArea1 = new ChartArea();
            chartArea1.Name = "BarChartArea";
            barChart.ChartAreas.Add(chartArea1);
            barChart.Dock = System.Windows.Forms.DockStyle.Fill;
            legend2.Name = "Legend3";
            barChart.Legends.Add(legend2);

            barChart.Series.Clear();
            barChart.BackColor = Color.LightGray;
            barChart.Palette = ChartColorPalette.Fire;
            barChart.ChartAreas[0].BackColor = Color.Transparent;
            barChart.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            barChart.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            Series series = new Series
            {
                Name = "series2",
                IsVisibleInLegend = false,
                ChartType = SeriesChartType.Column
            };
            barChart.Series.Add(series);

            // first bar
            series.Points.Add(70000);
            var p1 = series.Points[0];
            p1.Color = Color.Red;
            p1.AxisLabel = "Hiren Khirsaria";
            p1.LegendText = "Hiren Khirsaria";
            p1.Label = "70000";

            // write out a file
            SaveFileDialog dialog = new SaveFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                barChart.SaveImage(dialog.FileName, ChartImageFormat.Png);
            }
        }

        private void EvaluationChartForm_Load(object sender, EventArgs e)
        {

        }
    }
}
