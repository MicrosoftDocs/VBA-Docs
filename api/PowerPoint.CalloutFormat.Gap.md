---
title: CalloutFormat.Gap property (PowerPoint)
keywords: vbapp10.chm559013
f1_keywords:
- vbapp10.chm559013
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.Gap
ms.assetid: f7fa7ba7-18ab-2b72-a6a1-5bc252e47d49
ms.date: 06/08/2017
localization_priority: Normal
---


# CalloutFormat.Gap property (PowerPoint)

Returns or sets the horizontal distance (in points) between the end of the callout line and the text bounding box. Read/write.


## Syntax

_expression_.**Gap**

_expression_ A variable that represents a [CalloutFormat](PowerPoint.CalloutFormat.md) object.


## Return value

Single


## Example

This example sets the distance between the callout line and the text bounding box to 3 points for shape one on _myDocument_. For the example to work, shape one must be a callout.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).Callout.Gap = 3
```


## See also


[CalloutFormat Object](PowerPoint.CalloutFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]