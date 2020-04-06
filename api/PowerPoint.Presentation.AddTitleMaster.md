---
title: Presentation.AddTitleMaster method (PowerPoint)
keywords: vbapp10.chm583006
f1_keywords:
- vbapp10.chm583006
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.AddTitleMaster
ms.assetid: b49baa5b-217a-ab6d-3cb3-ff74e533ef20
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.AddTitleMaster method (PowerPoint)

Adds a title master to the specified presentation and returns a **[Master](PowerPoint.Master.md)** object that represents the title master.


## Syntax

_expression_. `AddTitleMaster`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

Master


## Remarks

If the presentation already has a title master, an error occurs.


## Example

This example adds a title master to the active presentation if it doesn't already have one.


```vb
With Application.ActivePresentation

    If Not .HasTitleMaster Then .AddTitleMaster

End With
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]