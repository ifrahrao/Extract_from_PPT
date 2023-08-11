import aspose.slides as slides

# Get all the text from presentation
text = slides.PresentationFactory().get_presentation_text("samplepptx.pptx", slides.TextExtractionArrangingMode.UNARRANGED)

# Load the presentation to get slide count
with slides.Presentation("samplepptx.pptx") as ppt:
    print("Total Slides",ppt.slides.length)

    # Loop through slides in the presentation
    for index in range(ppt.slides.length):

        # Print text of desired sections such as slide's text, layout text, notes, etc.
        print(text.slides_text[index].text)
        print(text.slides_text[index].layout_text)
        print(text.slides_text[index].master_text)
        print(text.slides_text[index].notes_text)