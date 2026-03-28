#slide 1
## layout
title
## slide title
Example Lecture
## subtitle
Minimal markdown input for the PPTX builder
## teleprompter
This example shows the smallest useful shape of a lecture deck source file.

#slide 2
## layout
bullet
## slide title
Supported Content
## bullets
- Real template placeholders are filled instead of drawing text boxes manually
- Presenter notes come from the teleprompter section
- Images can be inserted into picture placeholders
## teleprompter
This example uses the common title and bullets flow.

#slide 3
## layout
definition
## term
Template-first slide generation
## definition
The output deck inherits its look from the PowerPoint template rather than from hardcoded per-slide formatting.
## teleprompter
Definition and quote-like content can map to statement-style layouts.

#slide 4
## layout
photo
## slide title
Image Example
## subtitle
This slide resolves a local file path
## image
example-assets/sample-photo.png
## teleprompter
If the template layout contains a real picture placeholder, the builder inserts the image there.
