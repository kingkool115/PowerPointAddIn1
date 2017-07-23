using System;
using System.Windows.Forms;
using LiveCharts;
using LiveCharts.Wpf;

namespace PowerPointAddIn1
{
    public partial class EvaluationChartForm : Form
    {
        public EvaluationChartForm(String question)
        {
            InitializeComponent();

            Func<ChartPoint, string> labelPoint = chartPoint =>
                string.Format("{0} ({1:P})", chartPoint.Y, chartPoint.Participation);

            labelQuestion.Text = question;

            pieChart1.Series = new SeriesCollection
            {
                new PieSeries
                {
                    Title = "Maria",
                    Values = new ChartValues<double> {3},
                    PushOut = 15,
                    DataLabels = true,
                    LabelPoint = labelPoint
                },
                new PieSeries
                {
                    Title = "Charles",
                    Values = new ChartValues<double> {4},
                    DataLabels = true,
                    LabelPoint = labelPoint
                },
                new PieSeries
                {
                    Title = "Frida",
                    Values = new ChartValues<double> {6},
                    DataLabels = true,
                    LabelPoint = labelPoint
                },
                new PieSeries
                {
                    Title = "Frederic",
                    Values = new ChartValues<double> {2},
                    DataLabels = true,
                    LabelPoint = labelPoint
                }
            };

            pieChart1.LegendLocation = LegendLocation.Bottom;
        }

        private void EvaluationChartForm_Load(object sender, EventArgs e)
        {

        }
    }
}
