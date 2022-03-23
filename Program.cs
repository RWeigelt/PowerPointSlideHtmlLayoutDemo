using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPointSlideHtmlLayoutDemo;
//
// This code opens a PowerPoint presentation and looks at the first slide
// for shapes with texts in the format "{{something}}" (so-called
// "insertion points"). From there it creates an HTML file that contains
// these texts with proper layout and styling.
//
// The code also creates a PNG image file which contains anything else on
// the PowerPoint slide that is not an insertion point (e.g. images or
// other texts). This PNG file is used in the HTML file as a background.
//
// A note on naming: "insertion points" are different from "placeholders".
//
// - Placeholders are a concept built into PowerPoint. A placeholder is a
//   special shape that is placed and styled in Slide Master view.
//   Then you add content to it in Normal view.
//
// - Insertion points are a concept implemented by this code. An insertion
//   point is simply any shape with a text like "{{something}}".
//
// So, a placeholder *can* be an insertion point, but an insertion
// point does not necessarily have to be placeholder.
//
// IMPORTANT: The project is a proof-of-concept for how to turn a PowerPoint
// slide into something that could be used in a web browser. It is not
// intended to be a general purpose, ready-to-use solution. Instead it can
// be used as a starting point for your own code.

//
// For the demo, we'll use the PowerPoint presentation that is copied
// to the output directory during compilation.
//
var pptxFilePath = Path.Combine(AppContext.BaseDirectory, "Example.pptx");
var powerPoint = new Microsoft.Office.Interop.PowerPoint.Application();
var presentations = powerPoint.Presentations;
var presentation = presentations.Open(pptxFilePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

//
// The size of the PNG file to be created for the background is hard-coded
// to be rather small so I can use it directly in the blog post. Obviously,
// you can change it to your needs, but keep the aspect ratio in mind.
// (see next section).
//
const int pngWidth = 640;
const int pngHeight = 360;

//
// For translating the position and size of shapes in PowerPoint's
// coordinate system to pixels, we need the size of a slide.
//
// Note that for this demo, the aspect ratio of a slide in the
// example PowerPoint matches the aspect ratio of the output PNG
// file. For a more generic solution, a possible approach would
// be to look at the size of the slides first and calculate the size of
// the PNG using some width/height constraint.
//
var pageSetup = presentation.PageSetup;
var slideWidth = pageSetup.SlideWidth;
var slideHeight = pageSetup.SlideHeight;

//
// The example PowerPoint file contains only one slide
//
var slides = presentation.Slides;
var slide = slides[1]; // one-based!

var insertionPoints = SlideInsertionPointHelper.CollectInsertionPoints(slide);

var htmlTemplate = PowerPointSlideHtmlLayoutDemo.Properties.Resources.HtmlTemplate;
var htmlPageTextGenerator = new HtmlPageTextGenerator(htmlTemplate, pngWidth, pngHeight, slideWidth, slideHeight);
htmlPageTextGenerator.AddInsertionPoints(insertionPoints);
//
// Create the HTML page containing the insertion points with proper layout and styling.
// Anything else on the PowerPoint slide that was not an insertion point (including other
// texts) will be part of the background bitmap.
//
var myPicturesPathPath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
var pngFilePath = Path.Combine(myPicturesPathPath, "Background.png");
var htmlFilePath = Path.Combine(myPicturesPathPath, "HtmlPage.html");

var htmlText = htmlPageTextGenerator.GetHtmlText();
File.WriteAllText(htmlFilePath, htmlText);

SlideInsertionPointHelper.HideInsertionPoints(insertionPoints);
slide.Export(pngFilePath, "PNG", pngWidth, pngHeight);

//
// Cleanup
//
// https://support.microsoft.com/en-us/topic/office-application-does-not-exit-after-automation-from-visual-studio-net-client-96068fdb-7a84-ecf0-3b91-282fae81a618
foreach (var insertionPoint in insertionPoints)
{
    Marshal.FinalReleaseComObject(insertionPoint);
}
Array.Clear(insertionPoints);
Marshal.FinalReleaseComObject(pageSetup);
Marshal.FinalReleaseComObject(slide);
Marshal.FinalReleaseComObject(slides);
presentation.Close();
Marshal.FinalReleaseComObject(presentation);
Marshal.FinalReleaseComObject(presentations);
powerPoint.Quit();
Marshal.FinalReleaseComObject(powerPoint);
GC.Collect();
GC.WaitForPendingFinalizers();
