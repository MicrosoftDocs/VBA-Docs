---
title: TextStyles object (PowerPoint)
keywords: vbapp10.chm578000
f1_keywords:
- vbapp10.chm578000
ms.prod: powerpoint
api_name:
- PowerPoint.TextStyles
ms.assetid: 5c56df6d-8f37-ebe7-2955-c6c5de1ed771
ms.date: 06/08/2017
localization_priority: Normal
---


# TextStyles object (PowerPoint)

A collection of three text styles - title text, body text, and default text - each of which is represented by a  **[TextStyle](PowerPoint.TextStyle.md)** object.


## Remarks

Each text style contains a  **[TextFrame](PowerPoint.TextFrame.md)** object that describes how text is placed within the text bounding box, a **[Ruler](PowerPoint.Ruler.md)** object that contains tab stops and outline indent formatting information, and a **[TextStyleLevels](PowerPoint.TextStyleLevels.md)** collection that contains outline text formatting information.


## Example

Use  **TextStyles** (_index_), where _index_ is either **ppBodyStyle**, **ppDefaultStyle**, or **ppTitleStyle**, to return a single **TextStyle** object. This example sets the margins for the notes body area on all the notes pages in the active presentation.


```vb
With ActivePresentation.NotesMaster _
        .TextStyles(ppBodyStyle).TextFrame
    .MarginBottom = 50
    .MarginLeft = 50
    .MarginRight = 50
    .MarginTop = 50
End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]