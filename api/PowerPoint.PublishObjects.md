---
title: PublishObjects object (PowerPoint)
keywords: vbapp10.chm634000
f1_keywords:
- vbapp10.chm634000
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObjects
ms.assetid: 7f32fe4e-2345-ce6c-77c9-9fabdf9c5a23
ms.date: 06/08/2017
localization_priority: Normal
---


# PublishObjects object (PowerPoint)

A collection of  **[PublishObject](PowerPoint.PublishObject.md)** objects representing the set of complete or partial loaded presentations that are available for publishing to HTML.


## Remarks

You can specify the content and attributes of the published presentation by setting various properties of the  **PublishObject** object. For example, the [SourceType](PowerPoint.PublishObject.SourceType.md)property defines the portion of a loaded presentation to be published. The [RangeStart](PowerPoint.PublishObject.RangeStart.md)property and the [RangeEnd](PowerPoint.PublishObject.RangeEnd.md)property specify the range of slides to publish, and the [SpeakerNotes](PowerPoint.PublishObject.SpeakerNotes.md)property designates whether or not to publish the speaker's notes.

You cannot add to the  **PublishObjects** collection.


## Example

Use the  **PublishObjects** property to return the **PublishObjects** collection. This example publishes slides three through five of the active presentation to HTML. It names the published presentation Mallard.htm.


```vb
With ActivePresentation.PublishObjects(1)

    .FileName = "C:\Test\Mallard.htm"

    .SourceType = ppPublishSlideRange

    .RangeStart = 3

    .RangeEnd = 5

    .Publish

End With
```

Use  **Item** (_index_), where _index_ is always "1", to return the single **PublishObject** object for a loaded presentation. There can be only one **PublishObject** object for each loaded presentation.

This example defines the  **PublishObject** object to be the entire active presentation by setting the **SourceType** property to **ppPublishAll**.




```vb
ActivePresentation.PublishObjects.Item(1).SourceType = ppPublishAll
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]