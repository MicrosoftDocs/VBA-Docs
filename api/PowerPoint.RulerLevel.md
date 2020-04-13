---
title: RulerLevel object (PowerPoint)
keywords: vbapp10.chm572000
f1_keywords:
- vbapp10.chm572000
ms.prod: powerpoint
api_name:
- PowerPoint.RulerLevel
ms.assetid: 601fa2ef-8d8d-1e1d-e349-034d3c2842a5
ms.date: 06/08/2017
localization_priority: Normal
---


# RulerLevel object (PowerPoint)

Contains first-line indent and hanging indent information for an outline level. 


## Remarks

The **RulerLevel** object is a member of the **[RulerLevels](PowerPoint.RulerLevels.md)** collection. The **RulerLevels** collection contains a **RulerLevel** object for each of the five available outline levels.


## Example

Use  **RulerLevels** (_index_), where _index_ is the outline level, to return a single **RulerLevel** object. The following example sets the first-line indent and hanging indent for outline level one in body text on the slide master for the active presentation.


```vb
With ActivePresentation.SlideMaster _
        .TextStyles(ppBodyStyle).Ruler.Levels(1)
    .FirstMargin = 9
    .LeftMargin = 54
End With
```

The following example sets the first-line indent and hanging indent for outline level one in shape two on slide one in the active presentation.




```vb
With ActivePresentation.SlideMaster.Shapes(2) _
        .TextFrame.Ruler.Levels(1)
    .FirstMargin = 9
    .LeftMargin = 54
End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]