---
title: RulerLevel2 object (Office)
ms.prod: office
api_name:
- Office.RulerLevel2
ms.assetid: f1660a26-5990-9524-33f0-a2e3410160f3
ms.date: 01/23/2019
localization_priority: Normal
---


# RulerLevel2 object (Office)

Contains first-line indent and hanging indent information for an outline level.


## Remarks

The **RulerLevel2** object is a member of the **RulerLevels2** collection. The **RulerLevels2** collection contains a **RulerLevel2** object for each of the five available outline levels.


## Example

Use **RulerLevels2**(_index_), where _index_ is the outline level, to return a single **RulerLevel2** object. The following example sets the first-line indent and hanging indent for outline level one in body text on the slide master for the active presentation.


```vb
With ActivePresentation.SlideMaster _ 
 .TextStyles(ppBodyStyle).Ruler2.Levels(1) 
 .FirstMargin = 9 
 .LeftMargin = 54 
End With 

```


## See also

- [RulerLevel2 object members](overview/Library-Reference/rulerlevel2-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]