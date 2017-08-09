using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Charting = System.Windows.Forms.DataVisualization.Charting;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using PPt = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.utils
{
    class SessionController
    {
        // MyRibbon
        MyRibbon MyRibbon { get; set; }

        // List of all evaluated and displayed evaluations
        public List<Evaluation> CompletedEvaluations { get; set; }

        // Session id of current active presentation
        String SessionId { get; set; }

        // slide index during presentation
        int? SlideIndexPresentation { get; set; }

        // current slide id during presentation
        int? SlideIdPresentation { get; set; }

        // current slides
        Slides Slides { get; set;}

        // current presentation
        Presentation presentation { get; set;}

        // Define PowerPoint Application object
        PPt.Application pptApplication;

        /*
         * Constructor.
         */
        public SessionController(MyRibbon myRibbon, PPt.Application pptApplication)
        {
            MyRibbon = myRibbon;
            this.pptApplication = pptApplication;
            CompletedEvaluations = new List<Evaluation>();
            pptApplication.SlideShowNextClick += new EApplication_SlideShowNextClickEventHandler(nextSlideInSlideShow);
            pptApplication.SlideShowOnPrevious += new EApplication_SlideShowOnPreviousEventHandler(previousSlideInSlideShow);
        }

        /*
         * Start presentation in fullscreen mode.
         */
        public void startPresentation(bool fromBeginning, int slideIndexToStart, Presentation presentation, Slides slides)
        {
            this.presentation = presentation;
            this.Slides = slides;
            
            this.SessionId = generateSessionId();
            if (MyRibbon.myRestHelper == null)
            {
                MyRibbon.myRestHelper = new RestHelperLARS();
            }

            var slideShowSettings = presentation.SlideShowSettings;
            if (fromBeginning)
            {
                SlideIndexPresentation = 1;
                SlideIdPresentation = slides[SlideIndexPresentation].SlideID;
            }
            else
            {
                slideShowSettings.StartingSlide = slideIndexToStart;
                slideShowSettings.EndingSlide = presentation.Slides.Count;
            }
            slideShowSettings.Run();
            // TODO: wählen über welche Vorlesung/Kapitel die Session geht
            MyRibbon.myRestHelper.startPresentationSession(this.SessionId, 1, 6);
        }

        /*
         * Create a Random String as session key
         */
        private string generateSessionId()
        {
            int Size = 9;
            Random random = new Random();
            string input = "abcdefghijklmnopqrstuvwxyz0123456789";
            StringBuilder builder = new StringBuilder();
            char ch;
            for (int i = 0; i < Size; i++)
            {
                ch = input[random.Next(0, input.Length)];
                builder.Append(ch);
            }
            return builder.ToString();
        }

        /*
         * Is called whenever switching to next slide during a slide show.
         */
        public void nextSlideInSlideShow(SlideShowWindow ssw, Effect nEffect)
        {
            int currentSlideId = ssw.View.Slide.SlideID;
            int currentSlideIndex = ssw.View.Slide.SlideIndex;
            
            // TODO: generate session Id
            if (SlideIdPresentation != currentSlideId)
            {
                List<Question> questionsToPush = MyRibbon.getCustomSlideById(currentSlideId).PushQuestionList;
                pushQuestions(questionsToPush);
                evaluateAnswers(currentSlideId, SessionId, presentation);
            }

            // update slideId and slideIndex of current presentation
            SlideIndexPresentation = currentSlideIndex;
            SlideIdPresentation = currentSlideId;
        }

        /*
         * Is called whenever switching to previous slide during a slide show.
         */
        public void previousSlideInSlideShow(SlideShowWindow ssw)
        {
            SlideIndexPresentation -= 1;
        }

        public void pushQuestions(List<Question> questionsToPush)
        {
            //TODO: check if spend longer than 5 seconds on this slide

            foreach (var question in questionsToPush)
            {

            }
        }

        public void evaluateAnswers(int slideId, string sessionId, Microsoft.Office.Interop.PowerPoint.Presentation presentation)
        {
            //TODO: check if spend longer than 5 seconds on this slide

            // check if current slide has answers to evaluate
            List<String> questionIds = new List<String>();
            if (MyRibbon.getCustomSlideById(slideId) != null)
            {
                foreach (var question in MyRibbon.getCustomSlideById(slideId).EvaluationList)
                {
                    if (question.EvaluateSlideId != null)
                    {
                        questionIds.Add(question.ID);
                    }
                }
            }

            if (questionIds.Count == 0)
            {
                // no questions to evaluate
                return;
            }

            // make REST request to get evaluated Data
            var evaluationList = MyRibbon.myRestHelper.EvaluateAnswers(questionIds, sessionId);

            // make a chart out of the data
            if (evaluationList != null)
            {
                Charting.ChartArea chartArea1 = new Charting.ChartArea();
                Charting.Chart barChart = new Charting.Chart();

                barChart.ChartAreas.Add(chartArea1);
                barChart.Dock = DockStyle.Fill;

                barChart.Series.Clear();
                barChart.BackColor = Color.White;
                barChart.Palette = Charting.ChartColorPalette.Fire;
                barChart.ChartAreas[0].BackColor = Color.Transparent;
                barChart.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                barChart.ChartAreas[0].AxisY.MajorGrid.Enabled = false;

                Charting.Series series = new Charting.Series
                {
                    Name = "series2",
                    IsVisibleInLegend = false,
                    ChartType = Charting.SeriesChartType.Column
                };
                barChart.Series.Add(series);

                // iterate evaluations that should be displayed after current slide
                foreach (var evaluation in evaluationList)
                {
                    // iterate the answers for each evaluation
                    foreach (var answer in evaluation.Answers)
                    {
                        // create a bar that represents one possible answer for given question
                        series.Points.Add(answer.Value);
                        var p1 = series.Points[0];
                        p1.Color = Color.Red;
                        p1.AxisLabel = answer.Key;  // the answer
                        p1.LegendText = answer.Key;
                        p1.Label = answer.Value.ToString(); // number of people who gave that answer
                    }

                    // write out a file
                    // create a directory to store all diagramm pictures for that presentation 
                    String evaluationPicsDir = presentation.Path + "/presentaion_evaluation_" + sessionId + "_" + DateTime.Now.ToString("M/d/yyyy");
                    if (!Directory.Exists(evaluationPicsDir))
                    {
                        Directory.CreateDirectory(evaluationPicsDir);
                    }
                    var imageFile = evaluationPicsDir + "/diagramm_of_question_" + evaluation.QuestionId + ".png";
                    barChart.SaveImage(imageFile, Charting.ChartImageFormat.Png);


                    // create slides with that data
                    // Add slide to presentation
                    var slideIndexToShowEvaluation = presentation.Slides.FindBySlideID(slideId).SlideIndex + 1;
                    Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout =
                        presentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];
                    var slide = presentation.Slides.AddSlide(slideIndexToShowEvaluation, customLayout);

                    // add title to that slide
                    var objText = slide.Shapes[1].TextFrame.TextRange;
                    objText.Text = evaluation.Question;
                    objText.Font.Name = "Arial";
                    objText.Font.Size = 32;

                    string ImageFile2 = imageFile;
                    RectangleF rect = new RectangleF(50, 100, 600, 245);
                    slide.Shapes.AddPicture(ImageFile2, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, 100, 200);
                    CompletedEvaluations.Add(evaluation);
                }
            }
        }
    }
}
