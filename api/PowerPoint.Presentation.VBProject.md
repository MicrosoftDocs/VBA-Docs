---
title: Presentation.VBProject property (PowerPoint)
keywords: vbapp10.chm583022
f1_keywords:
- vbapp10.chm583022
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.VBProject
ms.assetid: 76713c8c-2263-7a5a-8133-726cc94bd73a
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.VBProject property (PowerPoint)

Returns a **VBProject** object that represents the individual Visual Basic project for the presentation. Read-only.


## Syntax

_expression_. `VBProject`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

VBProject


## Example

This example changes the name of the Visual Basic project for the active presentation.


```vb
ActivePresentation.VBProject.Name = "TestProject"
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]