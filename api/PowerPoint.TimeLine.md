---
title: TimeLine object (PowerPoint)
keywords: vbapp10.chm649000
f1_keywords:
- vbapp10.chm649000
ms.prod: powerpoint
api_name:
- PowerPoint.TimeLine
ms.assetid: 0b5a8863-8329-48d0-cb0b-3b34e87acb76
ms.date: 06/08/2017
localization_priority: Normal
---


# TimeLine object (PowerPoint)

Stores animation information for a  **Master**, **Slide**, or **SlideRange** object.


## Example

Use the [TimeLine](PowerPoint.Master.TimeLine.md)property of the  **[Master](PowerPoint.Master.md)**, **[Slide](PowerPoint.Slide.md)**, or **[SlideRange](PowerPoint.SlideRange.md)** object to return a **TimeLine** object.

The  **TimeLine** object's **[MainSequence](PowerPoint.TimeLine.MainSequence.md)** property gains access to the main animation sequence, while the **[InteractiveSequences](PowerPoint.TimeLine.InteractiveSequences.md)** property gains access to the collection of interactive animation sequences of a slide or slide range.

To reference a timeline object, use syntax similar to these code examples:




```vb
ActivePresentation.Slides(1).TimeLine.MainSequence

ActivePresentation.SlideMaster.TimeLine.InteractiveSequences

ActiveWindow.Selection.SlideRange.TimeLine.InteractiveSequences
```


## Properties



|Name|
|:-----|
|[Application](PowerPoint.TimeLine.Application.md)|
|[InteractiveSequences](PowerPoint.TimeLine.InteractiveSequences.md)|
|[MainSequence](PowerPoint.TimeLine.MainSequence.md)|
|[Parent](PowerPoint.TimeLine.Parent.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]