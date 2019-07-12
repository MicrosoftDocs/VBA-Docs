---
title: PublishObject.Publish method (PowerPoint)
keywords: vbapp10.chm635010
f1_keywords:
- vbapp10.chm635010
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObject.Publish
ms.assetid: 890382ef-8aec-466d-40f9-e2bae6dc558b
ms.date: 06/08/2017
localization_priority: Normal
---


# PublishObject.Publish method (PowerPoint)

Creates a Web presentation (HTML format) from any loaded presentation. You can view the published presentation in a web browser.


## Syntax

_expression_.**Publish**

_expression_ A variable that represents a [PublishObject](PowerPoint.PublishObject.md) object.


## Remarks

You can specify the content and attributes of the published presentation by setting various properties of the  **[PublishObject](PowerPoint.PublishObject.md)** object. For example, the **[SourceType](PowerPoint.PublishObject.SourceType.md)** property defines the portion of a loaded presentation to be published. The **[RangeStart](PowerPoint.PublishObject.RangeStart.md)** property and the **[RangeEnd](PowerPoint.PublishObject.RangeEnd.md)** property specify the range of slides to publish, and the **[SpeakerNotes](PowerPoint.PublishObject.SpeakerNotes.md)** property designates whether or not to publish the speaker's notes.


## Example

This example publishes slides three through five of the active presentation to HTML. It names the published presentation Mallard.htm.


```vb
With ActivePresentation.PublishObjects(1)

    .FileName = "C:\Test\Mallard.htm"

    .SourceType = ppPublishSlideRange

    .RangeStart = 3

    .RangeEnd = 5

    .Publish

End With
```


## See also


[PublishObject Object](PowerPoint.PublishObject.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]