---
title: Slides.Paste method (PowerPoint)
keywords: vbapp10.chm530008
f1_keywords:
- vbapp10.chm530008
ms.prod: powerpoint
api_name:
- PowerPoint.Slides.Paste
ms.assetid: 313027d1-6f8b-9964-f0bd-4ba33c973743
ms.date: 06/08/2017
localization_priority: Normal
---


# Slides.Paste method (PowerPoint)

Pastes the slides on the Clipboard into the  **Slides** collection for the presentation. Specify where you want to insert the slides with the **Index** argument. Returns a **[SlideRange](PowerPoint.SlideRange.md)** object that represents the pasted objects. Each pasted slide becomes a member of the specified **Slides** collection.


## Syntax

_expression_.**Paste** (_Index_)

_expression_ A variable that represents a [Slides](PowerPoint.Slides.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Integer**|The index number of the slide that the slides on the Clipboard are to be pasted before. If this argument is omitted, the slides on the Clipboard are pasted after the last slide in the presentation.|

## Return value

SlideRange


## Remarks

Use the  **[ViewType](PowerPoint.DocumentWindow.ViewType.md)** property to set the view for a window before pasting the Clipboard contents into it. The following table shows what you can paste into each view.



|**Into this view**|**You can paste the following from the Clipboard**|
|:-----|:-----|
|Slide view or notes page view|Shapes, text, or entire slides. If you paste a slide from the Clipboard, an image of the slide will be inserted onto the slide, master, or notes page as an embedded object. If one shape is selected, the pasted text will be appended to the shape's text; if text is selected, the pasted text will replace the selection; if anything else is selected, the pasted text will be placed in it is own text frame. Pasted shapes will be added to the top of the z-order and won't replace selected shapes.|
|Outline view|Text or entire slides. You cannot paste shapes into outline view. A pasted slide will be inserted before the slide that contains the cursor.|
|Slide sorter view|Entire slides. You cannot paste shapes or text into slide sorter view. A pasted slide will be inserted at the cursor or after the last slide selected in the presentation.|

## Example

This example cuts slides three and five from the Old Sales presentation and then inserts them before slide four in the active presentation.


```vb
Presentations("Old Sales").Slides.Range(Array(3, 5)).Cut

ActivePresentation.Slides.Paste 4
```


## See also


[Slides Object](PowerPoint.Slides.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
