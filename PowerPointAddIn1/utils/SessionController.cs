using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Charting = System.Windows.Forms.DataVisualization.Charting;
using Microsoft.Office.Interop.PowerPoint;
using PPt = Microsoft.Office.Interop.PowerPoint;
using Timers = System.Timers;

namespace PowerPointAddIn1.utils
{
    public class SessionController
    {
        // MyRibbon
        MyRibbon MyRibbon { get; set; }

        // ids of current seesion.
        int LectureId { get; set; }
        int ChapterId { get; set; }

        // List of all custom slides
        public List<CustomSlide> CustomSlides { get; set; }

        // List with all evaluation slide ids which were created during this session
        public Dictionary<Slide, Evaluation> EvaluationSlides { get; set; }

        // Session id of current active presentation
        String SessionId { get; set; }

        // slide index during presentation
        int? SlideIndexPresentation { get; set; }

        // current slide id during presentation
        int? SlideIdPresentation { get; set; }

        // current slides
        Slides Slides { get; set; }

        // current presentation
        Presentation presentation { get; set; }

        // Define PowerPoint Application object
        PPt.Application pptApplication;

        // This Thread is used to update evaluation slide, if user is on update slide.
        Timers.Timer updateThread;

        // This thread is used to push and evaluate questions only if user spents more than 5 seconds during a presentation.
        Timers.Timer observeTimeSpentOnSlideThread;

        // the time when a slide was reached.
        DateTime onSlideTime;

        // coordinates and dimensions for the question and diagram images on evaluation slides
        int DiagramImageX { get; } = 300;
        int DiagramImageY { get; } = 200;
        int DiagramImageWidth { get; } = 800;
        int DiagramImageHeight { get; } = 400;
        int QuestionimageX { get; } = 30;
        int QuestionImageY { get; } = 200;
        int QuestionImageHeight { get; } = 200;
        int QuestionImageWidth { get; } = 200;
        public int UpdateEvaluationInterval { get; } = 5000;
        public int TimeSpentOnSlideBeforePushing { get; } = 1000;

        /*
         * Constructor.
         */
        public SessionController(MyRibbon myRibbon, PPt.Application pptApplication)
        {
            MyRibbon = myRibbon;
            initCustomSlides();
            updateThread = new Timers.Timer();
            observeTimeSpentOnSlideThread = new Timers.Timer();
            EvaluationSlides = new Dictionary<Slide, Evaluation>();
            this.pptApplication = pptApplication;
            pptApplication.SlideShowNextSlide -= slideShowNextSlide;
            pptApplication.SlideShowNextSlide += new EApplication_SlideShowNextSlideEventHandler(slideShowNextSlide);
        }

        /*
         * Remove EventHandlers once the session is terminated.
         * So on next session they will not interfer with the other event handlers. 
         * **/
        public void removeEventHandlers()
        {
            pptApplication.SlideShowNextSlide -= slideShowNextSlide;
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
            this.LectureId = Int32.Parse(lectureId);
            this.ChapterId = Int32.Parse(chapterId);

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
                MyRibbon.myRestHelper.startPresentationSession(this.SessionId, LectureId, ChapterId);
                if (getCustomSlideById(SlideIdPresentation) != null)
                {
                    startObserveTimer((int)SlideIdPresentation);
                    //pushAndEvaluateThread(null, null, (int)SlideIdPresentation);
                }
                return;
            }
        }

        /*
         * What happens when next slide appears.
         */
        public void slideShowNextSlide(SlideShowWindow ssw)
        {
            onSlideTime = DateTime.Now;

            // cancel update thread when entering a new slide
            updateThread.Dispose();

            int currentSlideId = ssw.View.Slide.SlideID;
            int currentSlideIndex = ssw.View.Slide.SlideIndex;

            // check if we are currently on a evaluation slide -> the forloop is a workaround
            foreach (var evaluationSlide in EvaluationSlides)
            {
                if (evaluationSlide.Key.SlideID == currentSlideId)
                {
                    var evaluation = EvaluationSlides[evaluationSlide.Key];
                    startUpdateThread(evaluation, ssw);
                    //updateSlide(null, null, evaluation, ssw);
                }
            }

            if (SlideIdPresentation != currentSlideId && MyRibbon.getCustomSlideById(currentSlideId) != null)
            {
                //pushAndEvaluateThread(null, null, currentSlideId);
                startObserveTimer(currentSlideId);
            }

            // update slideId and slideIndex of current presentation
            SlideIndexPresentation = currentSlideIndex;
            SlideIdPresentation = currentSlideId;
        }

        /*
         * Start a thread to look for new incoming answers if user is currently on a evaluation slide.
         */
        private void startUpdateThread(Evaluation evaluation, SlideShowWindow ssw)
        {
            updateThread = new Timers.Timer();
            updateThread.Interval = UpdateEvaluationInterval;
            updateThread.Elapsed += new Timers.ElapsedEventHandler((sender, e) => updateSlide(sender, e, evaluation, ssw));
            updateThread.Start();
        }

        /*
         * Start a new timer instance to observe user how long he spents on a slide.
         */
        private void startObserveTimer(int currentSlideId)
        {
            observeTimeSpentOnSlideThread.Dispose();
            observeTimeSpentOnSlideThread = new Timers.Timer();
            observeTimeSpentOnSlideThread.Elapsed += new Timers.ElapsedEventHandler((sender, e) => pushAndEvaluateThread(sender, e, currentSlideId));
            observeTimeSpentOnSlideThread.AutoReset = false;
            observeTimeSpentOnSlideThread.Start();
        }

        /*
         * push and evaluate answers only if spent more than 5 seconds on slide.
         */
        private void pushAndEvaluateThread(object sender, Timers.ElapsedEventArgs e, int currentSlideId)
        {
            // if still the same slide after 5 seconds -> push and evaluate question
            System.Threading.Thread.Sleep(TimeSpentOnSlideBeforePushing + 1000);
            if ((DateTime.Now - onSlideTime).Seconds > TimeSpentOnSlideBeforePushing/1000 && SlideIdPresentation == currentSlideId)
            {
            }
            List<Question> questionsToPush = getCustomSlideById(currentSlideId).PushQuestionList;
            pushQuestions(questionsToPush);
            evaluateAnswers(currentSlideId, SessionId, presentation);
        }

        /*
         * Push questions to students.
         */
        public void pushQuestions(List<Question> questionsToPush)
        {
            foreach (var question in questionsToPush)
            {
                // push question only if it wasn't pushed yet
                if (!question.IsPushed)
                {
                    String lectureId = LectureId.ToString();
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

        /*
         * The next slide at this point is an evaluation slide, which is created in this method.
         */
        public void evaluateAnswers(int slideId, string sessionId, Presentation presentation)
        {
            // check if current slide has answers to evaluate
            List<String> questionIds = new List<String>();
            if (MyRibbon.getCustomSlideById(slideId) != null)
            {
                foreach (var question in getCustomSlideById(slideId).EvaluationList)
                {
                    if (question.EvaluateSlideId != null && !question.IsEvaluated && question.EvaluateSlideId == slideId)
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
                // iterate evaluations that should be displayed after current slide
                foreach (var evaluation in evaluationList)
                {

                    Charting.Chart barChart = createBarChart(evaluation.Answers);
                    
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
                    var evaluationSlide = presentation.Slides.AddSlide(slideIndexToShowEvaluation, customLayout);

                    // add title to that slide
                    var objText = evaluationSlide.Shapes[1].TextFrame.TextRange;
                    objText.Text = evaluation.Question;
                    objText.Font.Name = "Arial";
                    objText.Font.Size = 24;

                    // this first image is always centered into the center of the slide, no matter what coordinates you pass
                    // workaround: add an empty image first
                    // TODO: create a folder an put image there.
                    var filePathEmptyImage = "C:\\Users\\User\\Documents\\Visual Studio 2017\\Projects\\PowerPointAddIn1\\PowerPointAddIn1\\Resources\\empty_image.png";
                    evaluationSlide.Shapes.AddPicture2(filePathEmptyImage, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0);

                    // if question contains an image -> download it and add it to the evaluation slide
                    if (evaluation.ImageURL != null && evaluation.ImageURL.Length > 0)
                    {
                        String pathToQuestionImage = MyRibbon.myRestHelper.downloadQuestionImage(evaluationPicsDir + "/pic_of_question_" + Utils.generateRandomString(), evaluation.ImageURL);
                        evaluationSlide.Shapes.AddPicture(pathToQuestionImage, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, 30, 200, 200, 200);
                    }

                    // add diagramm image
                    evaluationSlide.Shapes.AddPicture2(pathToDiagramImage, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, DiagramImageX, DiagramImageY);
                    
                    // add EvaluationSlide to the list.
                    evaluation.pathToDiagramImage = pathToDiagramImage;
                    int hash = EvaluationSlides.GetHashCode();
                    EvaluationSlides.Add(evaluationSlide, evaluation);
                }
            }
        }

        /*
         * Update the current evaluation slide.
         */
        private void updateSlide(object sender, Timers.ElapsedEventArgs e, Evaluation evaluation, SlideShowWindow ssw)
        {
            // make REST request to get evaluated Data
            var newIncomingAnswers = MyRibbon.myRestHelper.GetAnswersForQuestion(evaluation.QuestionId, SessionId);

            // number of answer options raised
            if (newIncomingAnswers.Count > evaluation.Answers.Count)
            {
                evaluation.Answers = newIncomingAnswers;
                createNewDiagram(evaluation, ssw, newIncomingAnswers);
                return;
            }

            // the number of answers raised
            foreach (var newAnswer in newIncomingAnswers)
            {
                foreach (var evaluationAnswer in evaluation.Answers)
                {
                    // the number of Answers of that evaluation has changed -> update slide with new answer diagramm
                    if (evaluationAnswer.Key == newAnswer.Key && evaluationAnswer.Value != newAnswer.Value)
                    {
                        evaluation.Answers = newIncomingAnswers;
                        createNewDiagram(evaluation, ssw, newIncomingAnswers);
                        return;
                    }
                }
            }
        }

        /*
         * Creates a new evaluation diagram and replaces the existing image.
         */
        private void createNewDiagram(Evaluation evaluation, SlideShowWindow ssw, Dictionary<String, int> newAnswers)
        {
            // delete answer diagram of current slide
            ssw.View.Slide.Shapes[ssw.View.Slide.Shapes.Count].Delete();

            // create a new bar chart
            Charting.Chart barChart = createBarChart(newAnswers);

            // replace old diagram image with new one
            barChart.SaveImage(evaluation.pathToDiagramImage, Charting.ChartImageFormat.Png);

            // add diagramm image to current slide
            ssw.View.Slide.Shapes.AddPicture2(evaluation.pathToDiagramImage, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, DiagramImageX, DiagramImageY);
        }

        /*
         * Creates and returns a BarChart instance with legend. 
         */
        private Charting.Chart createBarChart(Dictionary<String, int> newAnswers) {
            Charting.ChartArea chartArea1 = new Charting.ChartArea();
            Charting.Chart barChart = new Charting.Chart();
            barChart.Width = DiagramImageWidth;
            barChart.Height = DiagramImageHeight;

            barChart.Font = new System.Drawing.Font("Arial", 40);

            barChart.ChartAreas.Add(chartArea1);
            barChart.Dock = DockStyle.Fill;

            barChart.Series.Clear();
            barChart.BackColor = Color.Transparent;
            //barChart.Palette = Charting.ChartColorPalette.Fire;

            barChart.ChartAreas[0].BackColor = Color.Transparent;
            barChart.ChartAreas[0].AxisX.Title = "Answers";
            barChart.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Arial", 15);
            barChart.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            barChart.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            
            Charting.Series series = new Charting.Series
            {
                Name = "series2",
                IsVisibleInLegend = false,
                ChartType = Charting.SeriesChartType.Column
            };
            barChart.Series.Add(series);

            // iterate the answers of the new received answers
            int counter = 0;
            Random rnd = new Random();

            var legend = new Charting.Legend();
            legend.LegendStyle = Charting.LegendStyle.Table;
            legend.TableStyle = Charting.LegendTableStyle.Wide;
            legend.IsEquallySpacedItems = true;
            legend.IsTextAutoFit = true;
            legend.BackColor = Color.White;
            legend.Font = new System.Drawing.Font("Arial", 10);
            legend.Docking = Charting.Docking.Bottom;
            legend.Alignment = StringAlignment.Center;

            foreach (var answer in newAnswers)
            {
                series.Points.Add(answer.Value);
                series.Font = new System.Drawing.Font("Arial", 15);
                var p1 = series.Points[counter];
                p1.IsVisibleInLegend = true;
                p1.Color = Color.FromArgb(rnd.Next(120), rnd.Next(120), rnd.Next(120));
                //p1.AxisLabel = answer.Key;  // the answer
                p1.LegendText = answer.Key;
                p1.Label = answer.Value.ToString(); // number of people who gave that answer
                legend.CustomItems.Add(new Charting.LegendItem(answer.Key, p1.Color, String.Empty));
                counter++;
            }

            barChart.Legends.Add(legend);
            return barChart;
        }
    }    
}
