---
title: Application.NewPresentation property (PowerPoint)
keywords: vbapp10.chm502049
f1_keywords:
- vbapp10.chm502049
ms.prod: powerpoint
api_name:
- PowerPoint.Application.NewPresentation
ms.assetid: 9685db30-9d73-19ad-432b-8d79b2d6ee50
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.NewPresentation property (PowerPoint)

Returns a **NewFile** object that represents a presentation listed on the **New Presentation** task pane. Read-only.


## Syntax

_expression_. `NewPresentation`

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Return value

NewFile


## Example

This example lists a presentation on the  **New Presentation** task pane at the bottom of the last section in the pane.


```vb
Sub CreateNewPresentationListItem()

    Application.NewPresentation.Add FileName:="C:\Presentation.ppt"

    Application.CommandBars("Task Pane").Visible = True

End Sub
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]