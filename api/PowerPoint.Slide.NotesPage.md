---
title: Slide.NotesPage property (PowerPoint)
keywords: vbapp10.chm531022
f1_keywords:
- vbapp10.chm531022
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.NotesPage
ms.assetid: 8d102704-1660-cc5f-6701-d7bc67b5924b
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.NotesPage property (PowerPoint)

Returns a  **[SlideRange](PowerPoint.SlideRange.md)** object that represents the notes pages for the specified slide or range of slides. Read-only.


## Syntax

_expression_. `NotesPage`

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


## Return value

SlideRange


## Remarks

The  **NotesPage** property returns the notes page for either a single slide or a range of slides and allows you to make changes only to those notes pages. If you want to make changes that affect all notes pages, use the **[NotesMaster](PowerPoint.Presentation.NotesMaster.md)** property to return the **Slide** object that represents the notes master.


## Example

This example sets the background fill for the notes page for slide one in the active presentation.


```vb
With ActivePresentation.Slides(1). NotesPage 
    .FollowMasterBackground = False 
    .Background.Fill.PresetGradient _ 
        msoGradientHorizontal, 1, msoGradientLateSunset 
End With
```


> [!NOTE] 
> The following properties and methods will fail if applied to a  **SlideRange** object that represents a notes page: **Copy** method, **Cut** method, **Delete** method, **Duplicate** method, **HeadersFooters** property, **Hyperlinks** property, **Layout** property, **PrintSteps** property, **SlideShowTransition** property.


## See also


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]