---
title: PublishObject object (PowerPoint)
keywords: vbapp10.chm635000
f1_keywords:
- vbapp10.chm635000
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObject
ms.assetid: 9419bec4-d2a6-6a2c-6400-4e2e270ff603
ms.date: 06/08/2017
localization_priority: Normal
---


# PublishObject object (PowerPoint)

Represents a complete or partial loaded presentation that is available for publishing to HTML. The  **PublishObject** object is a member of the **[PublishObjects](PowerPoint.PublishObjects.md)** collection.


## Remarks

You can specify the content and attributes of the published presentation by setting various properties of the  **PublishObject** object. For example, the [SourceType](PowerPoint.PublishObject.SourceType.md)property defines the portion of a loaded presentation to be published. The [RangeStart](PowerPoint.PublishObject.RangeStart.md)property and the [RangeEnd](PowerPoint.PublishObject.RangeEnd.md)property specify the range of slides to publish, and the [SpeakerNotes](PowerPoint.PublishObject.SpeakerNotes.md)property designates whether or not to publish the speaker's notes.


## Example

Use  **PublishObjects** (_index_), where _index_ is always "1", to return the single object for a loaded presentation. There can be only one **PublishObject** object for each loaded presentation. This example publishes slides three through five of presentation two to HTML. It names the published presentation Mallard.htm.


```vb
With Presentations(2).PublishObjects(1)

    .FileName = "C:\Test\Mallard.htm"

    .SourceType = ppPublishSlideRange

    .RangeStart = 3

    .RangeEnd = 5

    .Publish

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]