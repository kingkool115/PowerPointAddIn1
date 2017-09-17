using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Newtonsoft.Json;
// add PowerPoint namespace
using PPt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Xml;
using System.IO;
using Newtonsoft.Json.Linq;

namespace PowerPointAddIn1.utils
{
    public class PowerPointNavigator
    {
        MyRibbon myRibbon;

        // Define PowerPoint Application object
        public PPt.Application pptApplication;

        // Define Presentation object
        public Presentation presentation;

        // represents all slides which will push notifications to students
        public List<CustomSlide> customSlides = new List<CustomSlide>();

        // Define Slide collection
        public Slides slides;
        Slide slide;
        List<int> listOfSlideIds = new List<int>();

        // Slide count
        int slidescount;

        // slide index
        public int SlideIndex { get; set; }

        // current slideId
        public int SlideId { get; set; }

        // session id when a new ARS session is started
        public String ArsSessionId { get; set; }

        // list of errorMessages when push/evaluation order not correct.
        List<String> errorMessages = new List<String>();


        public PowerPointNavigator()
        {
            
            try
            {
                // Get Running PowerPoint Application object
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;
                this.pptApplication.SlideSelectionChanged += new PPt.EApplication_SlideSelectionChangedEventHandler(slideChanged);
                this.pptApplication.AfterPresentationOpen += new EApplication_AfterPresentationOpenEventHandler(afterPresentationOpened);
                this.pptApplication.PresentationCloseFinal += new EApplication_PresentationCloseFinalEventHandler(onPresentationClosed);
                this.pptApplication.PresentationSave += new EApplication_PresentationSaveEventHandler(saveCustomSlides);
                this.pptApplication.AfterNewPresentation += new EApplication_AfterNewPresentationEventHandler(onNewPresentationOpened);
            }
            catch
            {
                MessageBox.Show("Please Run PowerPoint Firstly", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
        }

        /*
         * When >1 presentation is open and you want to close one of them.
         * Then init presentation and slides variables with currently active presentation.
         */
        private void onNewPresentationOpened(Presentation pres)
        {
            afterPresentationOpened(pres);
        }

        /*
            * When >1 presentation is open and you want to close one of them.
            * Then init presentation and slides variables with currently active presentation.
            */
        private void onPresentationClosed(Presentation pres)
        {
            // Get Presentation Object
            presentation = pptApplication.ActivePresentation;

            // Get Slide collection object
            slides = presentation.Slides;
            
            // Get Slide count
            slidescount = slides.Count;
            try
            {
                // Get selected slide object in normal view
                slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
            }
            catch
            {
                // set first slide as selected slide
                if (slides.Count > 0)
                {
                    slide = slides[1];
                }
            }
        }
        
        /*
         * Is called when a presentation is opened.
         */
        private void afterPresentationOpened(Presentation pre)
        {
            myRibbon = Globals.Ribbons.Ribbon;

            // a presentation is already opened. 
            if (presentation != null)
            {
                DialogResult dialogResult = MessageBox.Show("Can not open a second presentation window with activated LARS Plugin.\nIt could cause unwanted side effects.",
                   "Can not open a second presentation", MessageBoxButtons.OK);
                if (dialogResult == DialogResult.OK)
                {
                    pre.Close();
                    return;
                }
            }

            if (pptApplication != null)
            {
                // Get Presentation Object
                try {
                    presentation = pptApplication.ActivePresentation;
                } catch {
                    return;
                }

                // load custom slides from saved instance
                var savedJsonCustomSlides = GetJsonContentFromRootElement("customSlides");
                if (savedJsonCustomSlides != null)
                {
                    var customSlides = JsonConvert.DeserializeObject<List<CustomSlide>>(savedJsonCustomSlides);
                    if (JsonConvert.DeserializeObject<List<CustomSlide>>(savedJsonCustomSlides) != null)
                    {
                        myRibbon.pptNavigator.customSlides = customSlides;
                    }
                }

                // load lecture from saved instance
                var savedJsonLecture = GetJsonContentFromRootElement("LectureForThisPresentation");
                if (savedJsonLecture != null)
                {
                    var savedLecture = JsonConvert.DeserializeObject<Lecture>(savedJsonLecture);
                    if (JsonConvert.DeserializeObject<Lecture>(savedJsonLecture) != null)
                    {
                        myRibbon.LectureForThisPresentation = savedLecture;
                        myRibbon.select_lecture_group.Label = "Lecture: " + myRibbon.LectureForThisPresentation.Name;
                    }
                }

                var name = presentation.FullName;

                // Get Slide collection object
                slides = presentation.Slides;
                foreach (Slide sld in slides)
                {
                    listOfSlideIds.Add(sld.SlideID);
                }

                // Get Slide count
                slidescount = slides.Count;
                // Get current selected slide 
                try
                {
                    // Get selected slide object in normal view
                    slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                }
                catch
                {
                    // set first slide as selected slide
                    if (slides.Count > 0)
                    {
                        slide = slides[1];
                    }
                }
            }
        }

        /*
         * Is called when presentation saved.
         */
        public void saveCustomSlides(Presentation pres)
        {
            string json = JsonConvert.SerializeObject(new { myRibbon.pptNavigator.customSlides, myRibbon.LectureForThisPresentation});
            var fakeXML = "<?xml version='1.0' standalone='no'?><root><" + pres.Name + "> " + json + "</" + pres.Name + "></root>";

            foreach (CustomXMLPart customXml in pres.CustomXMLParts)
            {
                // delete all custom xmls to avoid duplicates
                try
                {
                    customXml.Delete();
                }
                catch (COMException)
                {
                    Console.WriteLine("Can not delete this xml because it's necessary to run presentation.");
                }
            }

            // if not existing yet then add new customXMLPart
            pres.CustomXMLParts.Add(fakeXML);
            pres.Save();
        }

        /*
         * Get customSlides or LectureForThisPresentation as JSON from CustomXMLParts.
         */
        private string GetJsonContentFromRootElement(string jsonKey)
        {
            var customXmlParts = presentation.CustomXMLParts;
            foreach (CustomXMLPart customXmlPart in customXmlParts)
            {
                var xml = customXmlPart.XML;
                var xmlReader = XmlReader.Create(new StringReader(xml));
                while (xmlReader.Read())
                {
                    if (xmlReader.Name == presentation.Name)
                    {
                        var savedJson = xmlReader.ReadElementContentAsString();
                        JObject jObject = JObject.Parse(savedJson);
                        JToken myRibbonAttribute = jObject[jsonKey];
                        if (myRibbonAttribute != null)
                        {
                            var attributObj = myRibbonAttribute.ToString();
                            return attributObj;
                        }
                    }
                }
            }
            return null;
        }

        /*
         * Is called whenever a slide in powerpoint is changed.
         */
        private void slideChanged(SlideRange sr)
        {
            if (presentation == null)
            {
                return;
            }

            foreach (Slide sld in sr)
            {
                // custom slides were removed
                if (presentation.Slides.Count < slidescount)
                {
                    removeCustomSlides(listOfSlideIds, slides);
                }

                // custom slides were added
                if (presentation.Slides.Count > slidescount)
                {
                    addCustomSlides(slides, listOfSlideIds);
                }

                // update all attributes
                SlideIndex = sld.SlideIndex;
                SlideId = sld.SlideID;
                slide = slides[SlideIndex];
                slidescount = slides.Count;

                // TODO: maybe focus just one slide when more slides where added
            }

            // aktualisiere den index der verschobenen slides
            foreach (Slide sld in slides)
            {
                // aktualisiere den index der verschobenen slides
                myRibbon.pptNavigator.updateCustomSlideIndexIfSlideDraggedAndDropped(sld.SlideID, sld.SlideIndex);                
            }

            // if selectQuestionsFor or evaluateQuestionsForm were open while slides changed,
            // than update their listviews
            if (myRibbon.selectQuestionsForm != null)
            {
                myRibbon.selectQuestionsForm.updateQuestionsPerSlideListView();
            }
            if (myRibbon.evaluateQuestionsForm != null)
            {
                myRibbon.evaluateQuestionsForm.updateListViews();
            }

            myRibbon.updateRibbonQuestionEvaluationCounter(SlideId);
            checkPushEvaluationOrder();
        }

        /*
         * Check Push/evaluation order and show error Messages only if it wasn't shown before.
         */
        private void checkPushEvaluationOrder()
        {
            List<String> newErrorMessages = myRibbon.pptNavigator.checkQuestionsPushEvaluationOrder(false);

            // check if new error Messages are the same as before. if yes -> don't show dialog.
            if (newErrorMessages.Count == errorMessages.Count)
            {
                var sameErrorMessages = true;
                foreach (var errorMessage in errorMessages)
                {
                    if (!newErrorMessages.Contains(errorMessage))
                    {
                        sameErrorMessages = false;
                    }
                }

                if (sameErrorMessages)
                {
                    return;
                }
                errorMessages = newErrorMessages;
            }

            if (newErrorMessages.Count > errorMessages.Count)
            {
                String errMessage = "";
                foreach (var message in newErrorMessages)
                {
                    // show only the latest errorMessages
                    if (!errorMessages.Contains(message))
                    {
                        errMessage = errMessage + message + "\n\n";
                    }
                }
                DialogResult dialogResult = MessageBox.Show(errMessage + "\n\nDo you want to keep this slides order?",
                    "Wrong push/evaluation order.", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                {
                    myRibbon.pptNavigator.pptApplication.CommandBars.ExecuteMso("Undo");
                    errorMessages = new List<string>();
                }
                errorMessages = newErrorMessages;
            }

            if (newErrorMessages.Count < errorMessages.Count)
            {
                errorMessages = newErrorMessages;
            }
        }

        /*
         * Remove custom slides. 
         */
        public void removeCustomSlides(List<int> slideIdsBeforeRemove, Slides slidesAfterRemove)
        {
            // identify all removed slides
            List<int> removedSlides = new List<int>();
            foreach (int slideIdBefore in slideIdsBeforeRemove)
            {
                bool slideStillExisting = false;
                foreach (Slide slideAfterRemove in slidesAfterRemove)
                {
                    // slide is still existing
                    if (slideAfterRemove.SlideID == slideIdBefore)
                    {
                        slideStillExisting = true;
                        break;
                    }

                }

                if (!slideStillExisting)
                {
                    removedSlides.Add(slideIdBefore);
                }
            }

            // update slideIdsBeforeRemove
            foreach (int removedId in removedSlides)
            {
                slideIdsBeforeRemove.Remove(removedId);
            }

            // check if the removed Slides pushed questions. If yes -> delete the evaluations on the other slides.
            foreach (int removedSlideId in removedSlides)
            {
                var removedCustomSlide = myRibbon.pptNavigator.getCustomSlideById(removedSlideId);
                if (removedCustomSlide != null)
                {
                    foreach (Question question in removedCustomSlide.PushQuestionList)
                    {
                        var evaluationSlide = myRibbon.pptNavigator.getCustomSlideById(question.EvaluateSlideId);
                        evaluationSlide.EvaluationList.Remove(question);
                        evaluationSlide.removeEvaluation(question);
                        question.EvaluateSlideId = null;
                    }
                }
            }
        }

        /*
         * Remove custom slides. 
         */
        public void addCustomSlides(Slides slidesAfterAdding, List<int> slideIdsBeforeAdding)
        {
            // identify all removed slides
            List<int> addedSlides = new List<int>();
            foreach (Slide slideIdAfter in slidesAfterAdding)
            {
                bool isNewSlide = true;
                foreach (int slideIdBefore in slideIdsBeforeAdding)
                {
                    // slide is still existing
                    if (slideIdBefore == slideIdAfter.SlideID)
                    {
                        isNewSlide = false;
                        break;
                    }

                }

                if (isNewSlide)
                {
                    addedSlides.Add(slideIdAfter.SlideID);
                }
            }

            // update slideIdsBeforeAdding
            foreach (int addedId in addedSlides)
            {
                slideIdsBeforeAdding.Add(addedId);
            }
        }

        /*
         * Return a slide by given ID.
         */
        public Slide getSlideById(int? sldId)
        {
            foreach (Slide sld in slides)
            {
                if (sld.SlideID == sldId)
                {
                    return sld;
                }
            }
            return null;
        }

        /// <summary>
        /// Check if pushed questions will ever be evaluated, if they are pushed/evaluated in the given order.
        /// </summary>
        /// <param name="explicitCheck">true if check button is clicked</param>
        /// <returns></returns>
        public List<String> checkQuestionsPushEvaluationOrder(bool explicitCheck)
        {
            List<String> errorMessages = new List<String>();
            foreach (var customSlide in customSlides)
            {
                foreach (var question in customSlide.PushQuestionList)
                {
                    // question will never be evaluated because no slide is set to evaluate it
                    if (question.EvaluateSlideId == null)
                    {
                        // only add this error message when checking by clicked check button.
                        if (explicitCheck)
                        {
                            errorMessages.Add("Question pushed on slide number " + customSlide.SlideIndex + " will " +
                                "never be evaluated because you didn't define a slide to evaluate it.");
                        }
                        continue;
                    }

                    // question will never be evaluated because evaluationIndex <= pushIndex
                    if (getSlideById(question.EvaluateSlideId).SlideIndex <= question.PushSlideIndex)
                    {
                        errorMessages.Add("Question pushed on slide number " + question.PushSlideIndex + " will " +
                            "never be evaluated because you try to evaluate it on a previous or same slide number (slide number: " +
                            getSlideById(question.EvaluateSlideId).SlideIndex + ").");
                    }
                }
            }

            return errorMessages;
        }

        /*
         * Is called when Add-Evaluation-Button is clicked in EvaluateQuestionsForm.
         * Provide a slide index to EvaluateSlideIndex-attribute of a question.
         */
        public void addEvaluationToSlide(int slideIdToEvaluate, int slideIndex, Question question)
        {
            if (getCustomSlideById(slideIdToEvaluate) != null)
            {
                // find slide in questionSlides and add question
                getCustomSlideById(slideIdToEvaluate).addEvaluation(question);
            }
            else
            {
                // create new CustomSlide in questionSlides list
                customSlides.Add(new CustomSlide(slideIdToEvaluate, slideIndex, question, true));
            }
        }

        /*
         * Removes question from a certain slide.
         */
        public void removeEvaluationFromSlide(int slideId, Question question)
        {
            if (getCustomSlideById(slideId) != null)
            {
                getCustomSlideById(slideId).removeEvaluation(question);
            }
        }

        /*
         * Add question to a certain slide.
         */
        public void addQuestionToSlide(int slideId, int slideIndex, Question question)
        {
            if (getCustomSlideById(slideId) != null)
            {
                // find slide in questionSlides and add question
                getCustomSlideById(slideId).addPushQuestion(question);
            }
            else
            {
                // create new CustomSlide in questionSlides list
                customSlides.Add(new CustomSlide(slideId, slideIndex, question, false));
            }
        }

        /*
         * Removes question from a certain slide.
         */
        public void removeQuestionFromSlide(int slideId, Question question)
        {
            if (getCustomSlideById(slideId) != null)
            {
                getCustomSlideById(slideId).PushQuestionList.Remove(question);
                // remove the evaluation of that question as well
                getCustomSlideById(question.EvaluateSlideId).removeEvaluation(question);
            }
        }

        /*
         * Is called whenever slides are added/removed from presentation.
         */
        public void incrementDecrementCustomSlideIndexes(int position, int incDecSlideIndexValue)
        {
            foreach (var customSlide in customSlides)
            {
                // wenn es sich um einen Slide handelt, dessen index >= ist als das
                // hinzugefügte/gelöschte slide. Denn nur ist der Index des customSlide betroffen.
                if (customSlide.SlideIndex >= position)
                {
                    // new slide index is always higher than 0
                    if (customSlide.SlideIndex + incDecSlideIndexValue > 0)
                    {
                        customSlide.updatePushSlideIndex(customSlide.SlideIndex + incDecSlideIndexValue);
                    }
                }
            }
        }

        /*
         * Is called whenever slides are dragged & dropped.
         */
        public void updateCustomSlideIndexIfSlideDraggedAndDropped(int slideId, int newSlideIndex)
        {
            if (getCustomSlideById(slideId) != null)
            {
                getCustomSlideById(slideId).updatePushSlideIndex(newSlideIndex);
            }
        }

        /*
         * Check if a CustomSlide for given param slideIndex does already exist in questionSlides.
         */
        public CustomSlide getCustomSlideByIndex(int? slideIndex)
        {

            foreach (var slide in customSlides)
            {
                if (slide.SlideIndex == slideIndex)
                {
                    return slide;
                }
            }
            return null;
        }

        /*
         * Check if a CustomSlide for given param slideIndex does already exist in questionSlides.
         */
        public CustomSlide getCustomSlideById(int? slideId)
        {

            foreach (var slide in customSlides)
            {
                if (slide.SlideId == slideId)
                {
                    return slide;
                }
            }
            return null;
        }

        // Transform to First Page
        public void firstSlide()
        {
            try
            {
                // Call Select method to select first slide in normal view
                slides[1].Select();
                slide = slides[1];
            }
            catch
            {
                // Transform to first page in reading view
                pptApplication.SlideShowWindows[1].View.First();
                slide = pptApplication.SlideShowWindows[1].View.Slide;
            }
        }

        // Transform to Last Page
        public void lastSlide()
        {
            try
            {
                slides[slidescount].Select();
                slide = slides[slidescount];
            }
            catch
            {
                pptApplication.SlideShowWindows[1].View.Last();
                slide = pptApplication.SlideShowWindows[1].View.Slide;
            }
        }

        // Transform to next page
        public void nextSlide()
        {
            var slideIndexTmp = slide.SlideIndex + 1;
            if (slideIndexTmp > slidescount)
            {
                MessageBox.Show("It is already last page");
            }
            else
            {
                try
                {
                    slide = slides[slideIndexTmp];
                    slides[slideIndexTmp].Select();
                    // update current slideIndex
                    SlideIndex = slideIndexTmp;
                    SlideId = slide.SlideID;
                }
                catch
                {
                    pptApplication.SlideShowWindows[1].View.Next();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
        }

        // Transform to Last page
        public void previousSlide()
        {
            var slideIndexTmp = slide.SlideIndex - 1;
            if (slideIndexTmp >= 1)
            {
                try
                {
                    slide = slides[slideIndexTmp];
                    slides[slideIndexTmp].Select();
                    SlideIndex = slideIndexTmp;
                    SlideId = slide.SlideID;
                }
                catch
                {
                    pptApplication.SlideShowWindows[1].View.Previous();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
            else
            {
                MessageBox.Show("It is already Fist Page");
            }
        }
    }
}
