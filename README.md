# LightcycleProgressBar
This is a small Microsoft Visual Basic macro to add a TRON lightycle progress bar to the bottom of a PowerPoint presentation.  It works by drawing a lightcycle and jet wall on each slide with the jet wall image stretched to fit the trailing width.

To use the macro you will need to add both the `TRON-sml.png` and `trail-20x40.png` into the same directory as your presentation and also import the macro into PowerPoint from the `Module1.bas` file.

Once you are happy with your presentation **THEN**:
1. Make sure PowerPoint is in the *Normal* view (not *Slide Sorter*)
2. Click on one of the slides in your presentation
3. Run the `InsertSmallLightcycles` macro which will add a lightcycle and trail to each slide (or update existing lightcycles)
 
The lightcycle will start on the far left of the first slide and be just offscreen to the right of the last slide.

If you add slides, remove slides, or change the ordering then you will need to re-run the macro to update the progress bars.  If you notice any weird results with double lightcycles on some slides then simply running the macro again should clean it up.

Note that the macro cleans up existing cycles and trails on your slides by deleting any images it finds positioned 20 pixels up from the bottom of the slide.
This means that:
1. Any other images you have positioned on that line will also be deleted
2. If you move a cycle or trail so it's no longer on that line then it won't get cleaned up automagically if you re-run the script.

EOL
