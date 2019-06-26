---
title: HeadersFooters object (PowerPoint)
keywords: vbapp10.chm542000
f1_keywords:
- vbapp10.chm542000
ms.prod: powerpoint
api_name:
- PowerPoint.HeadersFooters
ms.assetid: 5fb10c90-0611-e797-836b-3f18b273af04
ms.date: 06/08/2017
localization_priority: Normal
---


# HeadersFooters object (PowerPoint)

Contains all the  **[HeaderFooter](PowerPoint.HeaderFooter.md)** objects on the specified slide, notes page, handout, or master.


## Remarks

Each  **HeaderFooter** object represents a header, footer, date and time, or slide number.


> [!NOTE] 
>  **HeaderFooter** objects aren't available for **[Slide](PowerPoint.Slide.md)** objects that represent notes pages. The **HeaderFooter** object that represents a header is available only for a notes master or handout master.


## Example

Use the  **[HeadersFooters](PowerPoint.Slide.HeadersFooters.md)** property to return the **HeadersFooters** object. Use the **[DateAndTime](PowerPoint.HeadersFooters.DateAndTime.md)**, **[Footer](PowerPoint.HeadersFooters.Footer.md)**, **[Header](PowerPoint.HeadersFooters.Header.md)**, or **[SlideNumber](PowerPoint.HeadersFooters.SlideNumber.md)** property to return an individual **HeaderFooter** object. The following example sets the footer text for slide one in the active presentation.


```vb
ActivePresentation.Slides(1).HeadersFooters.Footer _
    .Text = "Volcano Coffee"
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]