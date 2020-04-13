---
title: TextStyleLevel object (PowerPoint)
keywords: vbapp10.chm581000
f1_keywords:
- vbapp10.chm581000
ms.prod: powerpoint
api_name:
- PowerPoint.TextStyleLevel
ms.assetid: cf9a46d6-24f1-9679-4fe9-8c431d97ef92
ms.date: 06/08/2017
localization_priority: Normal
---


# TextStyleLevel object (PowerPoint)

Contains character and paragraph formatting information for an outline level. 


## Remarks

The **TextStyleLevel** object is a member of the **[TextStyleLevels](PowerPoint.TextStyleLevels.md)** collection. The **TextStyleLevels** collection contains one **TextStyleLevel** object for each of the five outline levels.


## Example

Use  **Levels** (_index_), where _index_ is a number from 1 through 5 that corresponds to the outline level, to return a single **TextStyleLevel** object. The following example sets the font name and font size, the space before paragraphs, and the paragraph alignment for level-one body text on all the slides in the active presentation.


```vb
With ActivePresentation.SlideMaster _
        .TextStyles(ppBodyStyle).Levels(1)
    With .Font
        .Name = "Arial"
        .Size = 36
    End With
    With .ParagraphFormat
        .LineRuleBefore = False
        .SpaceBefore = 14
        .Alignment = ppAlignJustify
    End With
End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]