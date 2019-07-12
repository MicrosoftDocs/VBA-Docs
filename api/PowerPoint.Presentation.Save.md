---
title: Presentation.Save method (PowerPoint)
keywords: vbapp10.chm583035
f1_keywords:
- vbapp10.chm583035
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Save
ms.assetid: 6d1251bb-27f3-0a80-bc2f-d385e2b3e3ec
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.Save method (PowerPoint)

Saves the specified presentation.


## Syntax

_expression_.**Save**

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Remarks

Use the  **[SaveAs](PowerPoint.Presentation.SaveAs.md)** method to save a presentation that has not been previously saved. To determine whether a presentation has been saved, test for a nonempty value for the **[FullName](PowerPoint.Presentation.FullName.md)** or **[Path](PowerPoint.Presentation.Path.md)** property. If a document that has the same name as the specified presentation already exists on disk, that document will be overwritten. No warning message is displayed.

To mark the presentation as saved without writing it to disk, set the  **Saved** property to **True**.


## Example

This example saves the active presentation if it is been changed since the last time it was saved.


```vb
With Application.ActivePresentation

    If Not .Saved And .Path <> "" Then .Save

End With


```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]