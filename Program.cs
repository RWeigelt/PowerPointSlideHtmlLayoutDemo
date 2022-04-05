using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
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
// So, a placeholder *can* be used for an insertion point, but an insertion
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

//
// The code may be running on a machine where the user has PowerPoint already
// running. In this case we do not want to close the app after we're done.
//
var powerPointAlreadyRunning = PowerPointHelper.TryGetRunningApplication(out var powerPoint);
if (!powerPointAlreadyRunning)
{
    powerPoint = new Microsoft.Office.Interop.PowerPoint.Application();
}
//
// We'll also check whether the presentation is already open and use it in its
// current state, regardless of whether it has unsaved changes. An alternative
// approach could be to always copy the PPTX file to a temporary location
// and open it there.
// In both cases, you could check for presentation.Saved== MsoTriState.msoTrue
// and, if this is not the case, ask the user if he/she would like to save the
// presentation (via presentation.Save());
//
var presentations = powerPoint.Presentations;
var presentationAlreadyOpen = PowerPointHelper.TryGetOpenPresentation(presentations, pptxFilePath, out var presentation);
if (!presentationAlreadyOpen)
{
    presentation = presentations.Open(pptxFilePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
}
    
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

var htmlText = htmlPageTextGenerator.GetHtmlText("url(Background.png)");
File.WriteAllText(htmlFilePath, htmlText);
//
// To create the PNG file, we'll have to hide all insertion points. 
//
if (presentationAlreadyOpen)
{
    // If the presentation has already been opened by the user, it should remain open and,
    // most importantly, remain unmodified after the creation of the PNG.
    // This is why we'll start a new undo entry, perform the operation and undo the actions.
    powerPoint.StartNewUndoEntry();
}
SlideInsertionPointHelper.HideInsertionPoints(insertionPoints);
slide.Export(pngFilePath, "PNG", pngWidth, pngHeight);
CommandBars commandBars = null;
if (presentationAlreadyOpen)
{
    commandBars = powerPoint.CommandBars;
    try
    {
        commandBars.ExecuteMso("Undo"); // This fails if no PowerPoint window is open
        powerPoint.StartNewUndoEntry();
    }
    catch
    {
        // just do nothing
    }
}
//
// Alternative to separate HTML and PNG files: Embed the PNG file into the HTML
// file so we have just one file per slide.
//
var bytes = File.ReadAllBytes(pngFilePath);
var imageData = Convert.ToBase64String(bytes);
var dataUrl = $"url('data:image/png;base64,{imageData}')";
htmlText = htmlPageTextGenerator.GetHtmlText(dataUrl);
var htmlFilePath2 = Path.Combine(myPicturesPathPath, "HtmlPage2.html");
File.WriteAllText(htmlFilePath2, htmlText);
//
// Cleanup
//
if (commandBars != null)
{
    Marshal.ReleaseComObject(commandBars);
}
foreach (var insertionPoint in insertionPoints)
{
    Marshal.ReleaseComObject(insertionPoint);
}
Array.Clear(insertionPoints);
Marshal.ReleaseComObject(pageSetup);
Marshal.ReleaseComObject(slide);
Marshal.ReleaseComObject(slides);
if (!presentationAlreadyOpen)
{
    presentation.Close();
}
Marshal.ReleaseComObject(presentation);
Marshal.ReleaseComObject(presentations);
if (!powerPointAlreadyRunning)
{
    powerPoint.Quit();
}
Marshal.ReleaseComObject(powerPoint);
GC.Collect();
GC.WaitForPendingFinalizers();
