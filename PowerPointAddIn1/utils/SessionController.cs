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
    public class SessionController
    {
        // MyRibbon
        MyRibbon MyRibbon { get; set; }
        
        // // List of all custom slides
        public List<CustomSlide> CustomSlides { get; set; }

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
            initCustomSlides();
            this.pptApplication = pptApplication;
            pptApplication.SlideShowNextClick -= nextSlideInSlideShow;
            pptApplication.SlideShowOnPrevious -= previousSlideInSlideShow;
            pptApplication.SlideShowNextClick += new EApplication_SlideShowNextClickEventHandler(nextSlideInSlideShow);
            pptApplication.SlideShowOnPrevious += new EApplication_SlideShowOnPreviousEventHandler(previousSlideInSlideShow);
        }

        /*
         * Remove EventHandlers once the session is terminated.
         * So on next session they will not interfer with the other event handlers. 
         * **/
        public void removeEventHandlers()
        {
            pptApplication.SlideShowNextClick -= nextSlideInSlideShow;
            pptApplication.SlideShowOnPrevious -= previousSlideInSlideShow;
        }
        
        /*
         * Init List of CustomSlides with completly new instances to avoid call-by-reference.
         * This is a workaroud. Could be done nicer.
         **/
        private void initCustomSlides()
        {
            CustomSlides = new List<CustomSlide>();
            foreach (var cs in MyRibbon.customSlides)
            {
                // iterate PushQuestionList
                List<Question> pushQuestionListOfNewCustomSlide = new List<Question>();
                foreach (var qu in cs.PushQuestionList)
                {
                    Question newQuestion = new Question(qu.Lecture, qu.Chapter, qu.Survey, qu.PushSlideId, qu.PushSlideIndex, qu.EvaluateSlideId, qu.EvaluateSlideIndex,
                                                        qu.ID, qu.isTextResponse, qu.IsPushed, qu.IsEvaluated);
                    pushQuestionListOfNewCustomSlide.Add(newQuestion);
                }

                // iterate PushQuestionList
                List<Question> evaluationQuestionListOfNewCustomSlide = new List<Question>();
                foreach (var qu in cs.EvaluationList)
                {
                    Question newQuestion = new Question(qu.Lecture, qu.Chapter, qu.Survey, qu.PushSlideId, qu.PushSlideIndex, qu.EvaluateSlideId, qu.EvaluateSlideIndex,
                                                        qu.ID, qu.isTextResponse, qu.IsPushed, qu.IsEvaluated);
                    evaluationQuestionListOfNewCustomSlide.Add(newQuestion);
                }
                CustomSlide newCustomSlide = new CustomSlide(cs.SlideId, cs.SlideIndex, pushQuestionListOfNewCustomSlide, evaluationQuestionListOfNewCustomSlide);

                // Add pushQuestions and EvaluationQuestions to new CustomSlide
                CustomSlides.Add(newCustomSlide);
            }
        }

        /*
         * Check if a CustomSlide for given param slideIndex does already exist in questionSlides.
         */
        public CustomSlide getCustomSlideById(int? slideId)
        {

            foreach (var slide in CustomSlides)
            {
                if (slide.SlideId == slideId)
                {
                    return slide;
                }
            }
            return null;
        }
        
        /*
         * Start presentation in fullscreen mode.
         */
        public void startPresentation(bool fromBeginning, int slideIndexToStart,
                                        Presentation presentation, Slides slides,
                                        String lectureId, String chapterId)
        {
            this.presentation = presentation;
            this.Slides = slides;
            
            this.SessionId = Utils.generateRandomString();
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
            if (lectureId != null && chapterId != null)
            {
                MyRibbon.myRestHelper.startPresentationSession(this.SessionId, Int32.Parse(lectureId), Int32.Parse(chapterId));
                return;
            }
            MyRibbon.myRestHelper.startPresentationSession(SessionId, null, null);
        }

        /*
         * Is called whenever switching to next slide during a slide show.
         */
        public void nextSlideInSlideShow(SlideShowWindow ssw, Effect nEffect)
        {
            int currentSlideId = ssw.View.Slide.SlideID;
            int currentSlideIndex = ssw.View.Slide.SlideIndex;
            
            if (SlideIdPresentation != currentSlideId && MyRibbon.getCustomSlideById(currentSlideId) != null)
            {
                List<Question> questionsToPush = getCustomSlideById(currentSlideId).PushQuestionList;
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
                // push question only if it wasn't pushed yet
                if (!question.IsPushed)
                {
                    String lectureId = question.Lecture.ID;
                    String questionId = question.ID;
                    String userEmail = MyRibbon.myRestHelper.userEmail;
                    var response = MyRibbon.myRestHelper.pushQuestion(questionId, lectureId, this.SessionId, userEmail);
                    if (response.StatusCode == System.Net.HttpStatusCode.OK)
                    {
                        question.IsPushed = true;
                    }
                }
            }
        }

        public void evaluateAnswers(int slideId, string sessionId, Presentation presentation)
        {
            //TODO: check if spend longer than 5 seconds on this slide

            // check if current slide has answers to evaluate
            List<String> questionIds = new List<String>();
            if (MyRibbon.getCustomSlideById(slideId) != null)
            {
                foreach (var question in getCustomSlideById(slideId).EvaluationList)
                {
                    if (question.EvaluateSlideId != null && !question.IsEvaluated)
                    {
                        questionIds.Add(question.ID);
                        question.IsEvaluated = true;
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
                    String pathToDiagramImage = evaluationPicsDir + "/diagramm_of_question_" + Utils.generateRandomString() + ".png";
                    barChart.SaveImage(pathToDiagramImage, Charting.ChartImageFormat.Png);


                    // create slides with that data
                    // Add slide to presentation
                    var slideIndexToShowEvaluation = presentation.Slides.FindBySlideID(slideId).SlideIndex + 1;
                    CustomLayout customLayout =
                        presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];
                    var slide = presentation.Slides.AddSlide(slideIndexToShowEvaluation, customLayout);

                    // add title to that slide
                    var objText = slide.Shapes[1].TextFrame.TextRange;
                    objText.Text = evaluation.Question;
                    objText.Font.Name = "Arial";
                    objText.Font.Size = 24;

                    // this first image is always centered into the center of the slide, no matter what coordinates you pass
                    // workaround: add an empty image first
                    // TODO: create a folder an put image there.
                    var filePathEmptyImage = "C:\\Users\\User\\Documents\\Visual Studio 2017\\Projects\\PowerPointAddIn1\\PowerPointAddIn1\\Resources\\empty_image.png";
                    slide.Shapes.AddPicture2(filePathEmptyImage, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0);

                    // if question contains an image -> download it and add it to the evaluation slide
                    if (evaluation.ImageURL != null && evaluation.ImageURL.Length > 0)
                    {
                        String pathToQuestionImage = MyRibbon.myRestHelper.downloadQuestionImage(evaluationPicsDir + "/pic_of_question_" + Utils.generateRandomString(), evaluation.ImageURL);
                        slide.Shapes.AddPicture(pathToQuestionImage, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, 30, 200, 200, 200);
                    }

                    // add diagramm image
                    slide.Shapes.AddPicture2(pathToDiagramImage, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, 300, 200);
                }
            }
        }
    }
}
