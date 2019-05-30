---
title: LayoutGuides object (Publisher)
keywords: vbapb10.chm1179647
f1_keywords:
- vbapb10.chm1179647
ms.prod: publisher
api_name:
- Publisher.LayoutGuides
ms.assetid: 7430c1c4-c7f5-d9b6-cea8-b21fe9e2905f
ms.date: 05/31/2019
localization_priority: Normal
---


# LayoutGuides object (Publisher)

Represents the measurement grid that appears superimposed on publication pages as an aid to laying out design elements.
 
## Remarks

Use the **[LayoutGuides](Publisher.Document.LayoutGuides.md)** property of the **Document** object to return a **LayoutGuides** object. 

Use the **LayoutGuide** object's margin properties and **Rows** and **Columns** properties to set how many rows and columns are displayed in the layout guides and where they appear on a page.

## Example

This example sets the margins of the active presentation to two inches.

```vb
With ActiveDocument.LayoutGuides 
 .MarginTop = Application.InchesToPoints(Value:=2) 
 .MarginBottom = Application.InchesToPoints(Value:=2) 
 .MarginLeft = Application.InchesToPoints(Value:=2) 
 .MarginRight = Application.InchesToPoints(Value:=2) 
End With
```


## Properties

- [Application](Publisher.LayoutGuides.Application.md)
- [ColumnGutterWidth](Publisher.LayoutGuides.ColumnGutterWidth.md)
- [Columns](Publisher.LayoutGuides.Columns.md)
- [GutterCenterlines](Publisher.LayoutGuides.GutterCenterlines.md)
- [HorizontalBaseLineOffset](Publisher.LayoutGuides.HorizontalBaseLineOffset.md)
- [HorizontalBaseLineSpacing](Publisher.LayoutGuides.HorizontalBaseLineSpacing.md)
- [MarginBottom](Publisher.LayoutGuides.MarginBottom.md)
- [MarginLeft](Publisher.LayoutGuides.MarginLeft.md)
- [MarginRight](Publisher.LayoutGuides.MarginRight.md)
- [MarginTop](Publisher.LayoutGuides.MarginTop.md)
- [MirrorGuides](Publisher.LayoutGuides.MirrorGuides.md)
- [Parent](Publisher.LayoutGuides.Parent.md)
- [RowGutterWidth](Publisher.LayoutGuides.RowGutterWidth.md)
- [Rows](Publisher.LayoutGuides.Rows.md)
- [VerticalBaseLineOffset](Publisher.LayoutGuides.VerticalBaseLineOffset.md)
- [VerticalBaseLineSpacing](Publisher.LayoutGuides.VerticalBaseLineSpacing.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]