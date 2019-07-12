---
title: Slide.Design property (PowerPoint)
keywords: vbapp10.chm531029
f1_keywords:
- vbapp10.chm531029
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Design
ms.assetid: bac64534-92f7-5611-db7e-501504e577e1
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.Design property (PowerPoint)

Returns a  **Design** object representing a design.


## Syntax

_expression_. `Design`

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


## Return value

Design


## Example

The following example adds a title master.


```vb
Sub AddDesignMaster

    ActivePresentation.Slides(1).Design.AddTitleMaster

End Sub
```


## See also


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]