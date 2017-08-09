using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        SessionController sessionController;

        // Define PowerPoint Application object
        PPt.Application pptApplication;

        // Define Presentation object
        public Presentation presentation;

        // Define Slide collection
        Slides slides;
        Slide slide;

        // Slide count
        int slidescount;

        // slide index
        public int SlideIndex { get; set; }

        // current slideId
        public int SlideId { get; set; }

        // session id when a new ARS session is started
        public String ArsSessionId { get; set; }
        
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
            }
            catch
            {
                MessageBox.Show("Please Run PowerPoint Firstly", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
        }
        
        /*
         * Start presentation in fullscreen mode.
         */
        public void startPresentation(bool fromBeginning)
        {
            if (fromBeginning)
            {
                sessionController.startPresentation(fromBeginning, 1, presentation, slides);
            } else {
                sessionController.startPresentation(fromBeginning, SlideIndex, presentation, slides);
            }
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
            sessionController = new SessionController(myRibbon, pptApplication);

            if (pptApplication != null)
            {
                // Get Presentation Object
                presentation = pptApplication.ActivePresentation;
                // load saved instace
                var savedJson = GetJsonContentFromRootElement();
                if (savedJson != null)
                {
                    var customSlides = JsonConvert.DeserializeObject<List<CustomSlide>>(savedJson);
                    if (JsonConvert.DeserializeObject<List<CustomSlide>>(savedJson) != null)
                    {
                        myRibbon.customSlides = customSlides;
                    }
                }

                var name = presentation.FullName;

                // Get Slide collection object
                slides = presentation.Slides;
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
            string json = JsonConvert.SerializeObject(new { myRibbon.customSlides });
            //json = @"{'?xml': {'@version': '1.0', '@standalone': 'no'}, 'root': " + json + "}";
            //XmlDocument doc = (XmlDocument)JsonConvert.DeserializeXmlNode(json);
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
         * Get customSlides as JSON from CustomXMLParts.
         */
        private string GetJsonContentFromRootElement()
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
                        JToken questionSlides = jObject["questionSlides"];
                        var slides = questionSlides.ToString();
                        return slides;
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

                if (presentation.Slides.Count < slidescount)
                {
                    myRibbon.removeCustomSlide(SlideId);
                }
                // TODO: wenn eingefügte slides Fragen habe, dann auch diese berücksichtigen

                // update alle attributes
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
                myRibbon.updateCustomSlideIndexIfSlideDraggedAndDropped(sld.SlideID, sld.SlideIndex);                
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
            myRibbon.checkQuestionsPushEvaluationOrder(false);
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
