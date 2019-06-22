---
title: Application.VBE property (PowerPoint)
keywords: vbapp10.chm502020
f1_keywords:
- vbapp10.chm502020
ms.prod: powerpoint
api_name:
- PowerPoint.Application.VBE
ms.assetid: 33a3d113-31f6-3705-cdb9-d5e07fa82820
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.VBE property (PowerPoint)

Returns a  **VBE** object that represents the Visual Basic Editor. Read-only.


## Syntax

_expression_.**VBE**

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Return value

VBE


## Example

This example sets the name of the active project in the Visual Basic Editor.


```vb
Application.VBE.ActiveVBProject.Name = "TestProject"
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]