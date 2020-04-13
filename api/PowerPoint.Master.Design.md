---
title: Master.Design property (PowerPoint)
keywords: vbapp10.chm533014
f1_keywords:
- vbapp10.chm533014
ms.prod: powerpoint
api_name:
- PowerPoint.Master.Design
ms.assetid: 78035fbd-e2f3-9089-2263-c04ce72394db
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.Design property (PowerPoint)

Returns a **Design** object representing a design.


## Syntax

_expression_. `Design`

_expression_ A variable that represents a [Master](PowerPoint.Master.md) object.


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


[Master Object](PowerPoint.Master.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]