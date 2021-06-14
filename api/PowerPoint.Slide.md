---
title: Slide object (PowerPoint)
keywords: vbapp10.chm535000
f1_keywords:
- vbapp10.chm535000
ms.prod: powerpoint
api_name:
- PowerPoint.Slide
ms.assetid: afe42344-6898-00d2-ecc1-b0ed23a71fe8
ms.date: 06/08/2019
localization_priority: Normal
---

# Slide object (PowerPoint)

Represents a slide. The **[Slides](PowerPoint.Slides.md)** collection contains all the **Slide** objects in a presentation.

## Remarks

> [!NOTE] 
> Don't be confused if you are trying to return a reference to a single slide but you end up with a **[SlideRange](PowerPoint.SlideRange.md)** object. A single slide can be represented either by a **Slide** object or by a [SlideRange](PowerPoint.SlideRange.md)collection that contains only one slide, depending on how you return a reference to the slide. For example, if you create and return a reference to a slide by using the **[Add](PowerPoint.Presentations.Add.md)** method, the slide is represented by a **Slide** object. However, if you create and return a reference to a slide by using the **[Duplicate](PowerPoint.Slide.Duplicate.md)** method, the slide is represented by a **SlideRange** collection that contains a single slide. Because all the properties and methods that apply to a **Slide** object also apply to a **SlideRange** collection that contains a single slide, you can work with the returned slide in the same way, regardless of whether it is represented by a **Slide** object or a **SlideRange** collection.


The following examples describe how to:

- Return a slide that you specify by name, index number, or slide ID number
    
- Return a slide in the selection
    
- Return the slide that's currently displayed in any document window or slide show window you specify
    
- Create a new slide
    

## Example

Use  **Slides** (_index_), where _index_ is the slide name or index number, or use **Slides.FindBySlideID** (_index_), where _index_ is the slide ID number, to return a single **Slide** object. The following example sets the layout for slide one in the active presentation.

```vb
ActivePresentation.Slides(1).Layout = ppLayoutTitle
```

The following example sets the layout for the slide with the ID number 265.


```vb
ActivePresentation.Slides.FindBySlideID(265).Layout = ppLayoutTitle
```

Use  **Selection.SlideRange** (_index_), where _index_ is the slide name or index number within the selection, to return a single **Slide** object. The following example sets the layout for slide one in the selection in the active window, assuming that there's at least one slide selected.

```vb
ActiveWindow.Selection.SlideRange(1).Layout = ppLayoutTitle
```

If there's only one slide selected, you can use  **Selection.SlideRange** to return a **SlideRange** collection that contains the selected slide. The following example sets the layout for slide one in the current selection in the active window, assuming that there's exactly one slide selected.

```vb
ActiveWindow.Selection.SlideRange.Layout = ppLayoutTitle
```

Use the  **Slide** property to return the slide that's currently displayed in the specified document window or slide show window view. The following example copies the slide that's currently displayed in document window two to the Clipboard.

```vb
Windows(2).View.Slide.Copy
```

Use the  **Add** method to create a new slide and add it to the presentation. The following example adds a title slide to the beginning of the active presentation.

```vb
ActivePresentation.Slides.Add 1, ppLayoutTitleOnly
```

## Methods
|Name|
|:-----|
|[ApplyTemplate](PowerPoint.Slide.ApplyTemplate.md)|
|[ApplyTemplate2](PowerPoint.slide.applytemplate2.md)|
|[ApplyTheme](PowerPoint.Slide.ApplyTheme.md)|
|[ApplyThemeColorScheme](PowerPoint.Slide.ApplyThemeColorScheme.md)|
|[Copy](PowerPoint.Slide.Copy.md)|
|[Cut](PowerPoint.Slide.Cut.md)|
|[Delete](PowerPoint.Slide.Delete.md)|
|[Duplicate](PowerPoint.Slide.Duplicate.md)|
|[Export](PowerPoint.Slide.Export.md)|
|[MoveTo](PowerPoint.Slide.MoveTo.md)|
|[MoveToSectionStart](PowerPoint.Slide.MoveToSectionStart.md)|
|[PublishSlides](PowerPoint.Slide.PublishSlides.md)|
|[Select](PowerPoint.Slide.Select.md)|

## Properties
|Name|
|:-----|
|[Application](PowerPoint.Slide.Application.md)|
|[Background](PowerPoint.Slide.Background.md)|
|[BackgroundStyle](PowerPoint.Slide.BackgroundStyle.md)|
|[ColorScheme](PowerPoint.Slide.ColorScheme.md)|
|[Comments](PowerPoint.Slide.Comments.md)|
|[CustomerData](PowerPoint.Slide.CustomerData.md)|
|[CustomLayout](PowerPoint.Slide.CustomLayout.md)|
|[Design](PowerPoint.Slide.Design.md)|
|[DisplayMasterShapes](PowerPoint.Slide.DisplayMasterShapes.md)|
|[FollowMasterBackground](PowerPoint.Slide.FollowMasterBackground.md)|
|[HasNotesPage](PowerPoint.Slide.HasNotesPage.md)|
|[HeadersFooters](PowerPoint.Slide.HeadersFooters.md)|
|[Hyperlinks](PowerPoint.Slide.Hyperlinks.md)|
|[Layout](PowerPoint.Slide.Layout.md)|
|[Master](PowerPoint.Slide.Master.md)|
|[Name](PowerPoint.Slide.Name.md)|
|[NotesPage](PowerPoint.Slide.NotesPage.md)|
|[Parent](PowerPoint.Slide.Parent.md)|
|[PrintSteps](PowerPoint.Slide.PrintSteps.md)|
|[sectionIndex](PowerPoint.Slide.sectionIndex.md)|
|[Shapes](PowerPoint.Slide.Shapes.md)|
|[SlideID](PowerPoint.Slide.SlideID.md)|
|[SlideIndex](PowerPoint.Slide.SlideIndex.md)|
|[SlideNumber](PowerPoint.Slide.SlideNumber.md)|
|[SlideShowTransition](PowerPoint.Slide.SlideShowTransition.md)|
|[Tags](PowerPoint.Slide.Tags.md)|
|[ThemeColorScheme](PowerPoint.Slide.ThemeColorScheme.md)|
|[TimeLine](PowerPoint.Slide.TimeLine.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
