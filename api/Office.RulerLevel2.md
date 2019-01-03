---
title: RulerLevel2 object (Office)
ms.prod: office
api_name:
- Office.RulerLevel2
ms.assetid: f1660a26-5990-9524-33f0-a2e3410160f3
ms.date: 06/08/2017
---


# RulerLevel2 object (Office)

Contains first-line indent and hanging indent information for an outline level.


## Remarks

The  **RulerLevel2** object is a member of the **RulerLevels2** collection. The **RulerLevels2** collection contains a **RulerLevel2** object for each of the five available outline levels.


## Example

Use  `RulerLevels2(index)`, where index is the outline level, to return a single  **RulerLevel2** object. The following example sets the first-line indent and hanging indent for outline level one in body text on the slide master for the active presentation.


```vb
With ActivePresentation.SlideMaster _ 
 .TextStyles(ppBodyStyle).Ruler2.Levels(1) 
 .FirstMargin = 9 
 .LeftMargin = 54 
End With 

```


## Properties



|**Name**|
|:-----|
|[Application](Office.RulerLevel2.Application.md)|
|[Creator](Office.RulerLevel2.Creator.md)|
|[FirstMargin](Office.RulerLevel2.FirstMargin.md)|
|[LeftMargin](Office.RulerLevel2.LeftMargin.md)|
|[Parent](Office.RulerLevel2.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
