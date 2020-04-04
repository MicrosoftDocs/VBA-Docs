---
title: Application.ActivePresentation property (PowerPoint)
keywords: vbapp10.chm503001
f1_keywords:
- vbapp10.chm503001
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ActivePresentation
ms.assetid: 55ff4906-09e5-2c5c-0ed7-5f7a767542f7
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ActivePresentation property (PowerPoint)

Returns a **[Presentation](PowerPoint.Presentation.md)** object that represents the presentation open in the active window. Read-only.


## Syntax

_expression_. `ActivePresentation`

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Return value

Presentation


## Remarks

 If an embedded presentation is in-place active, the **ActivePresentation** property returns the embedded presentation.


## Example

This example saves the loaded presentation to the application folder in a file named "TestFile."


```vb
MyPath = Application.Path & "\TestFile"

Application.ActivePresentation.SaveAs MyPath
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
