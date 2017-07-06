using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
// add PowerPoint namespace
using PPt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.utils
{
    public class PowerPointNavigator
    {
        MyRibbon myRibbon;
        // Define PowerPoint Application object
        PPt.Application pptApplication;

        // Define Presentation object
        public PPt.Presentation presentation;
       
        // Define Slide collection
        PPt.Slides slides;
        PPt.Slide slide;

        // Slide count
        int slidescount;
        // slide index
        int slideIndex;

        public PowerPointNavigator()
        {
            
            try
            {
                // Get Running PowerPoint Application object
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;
                this.pptApplication.SlideSelectionChanged += new PPt.EApplication_SlideSelectionChangedEventHandler(slideChanged);
                this.pptApplication.AfterPresentationOpen += new PPt.EApplication_AfterPresentationOpenEventHandler(afterPresentationOpened);
            }
            catch
            {
                MessageBox.Show("Please Run PowerPoint Firstly", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
        }

        private void afterPresentationOpened(Presentation pre)
        {
            myRibbon = Globals.Ribbons.Ribbon;
            if (pptApplication != null)
            {
                // Get Presentation Object
                presentation = pptApplication.ActivePresentation;
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
         * returns current slide index.
         */
        public int getCurrentSlide()
        {
            return slideIndex;
        }

        /*
         * Is called whenever a slide in powerpoint is changed.
         */
        private void slideChanged(SlideRange sr)
        {
            slideIndex = sr.SlideIndex;
            slide = slides[slideIndex];
            if (myRibbon.selectQuestionsForm != null)
            {
                myRibbon.selectQuestionsForm.updateQuestionsPerSlideListView();
            }
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
                    slideIndex = slideIndexTmp;
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
                    slideIndex = slideIndexTmp;
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
