---
title: Table object (PowerPoint)
keywords: vbapp10.chm622000
f1_keywords:
- vbapp10.chm622000
ms.prod: powerpoint
api_name:
- PowerPoint.Table
ms.assetid: ebbbca9f-4591-10ce-3c74-33b46a3b7cdf
ms.date: 06/08/2017
localization_priority: Normal
---


# Table object (PowerPoint)

Represents a table shape on a slide. The **Table** object is a member of the **Shapes** collection. The **Table** object contains the **[Columns](PowerPoint.Columns.md)** collection and the **[Rows](PowerPoint.Rows.md)** collection.


## Example

Use  **Shapes** (_index_), where _index_ is a number, to return a shape containing a table. Use the [HasTable](PowerPoint.Shape.HasTable.md)property to see if a shape contains a table. This example walks through the shapes on slide one, checks to see if each shape has a table, and then sets the mouse click action for each table shape to advance to the next slide.


```vb
With ActivePresentation.Slides(2).Shapes

    For i = 1 To .Count

        If .Item(i).HasTable Then

            .Item(i).ActionSettings(ppMouseClick) _

                .Action = ppActionNextSlide

        End If

    Next

End With
```

Use the [Cell](PowerPoint.Table.Cell.md)method of the  **Table** object to access the contents of each cell. This example inserts the text "Cell 1" in the first cell of the table in shape five on slide three.




```vb
ActivePresentation.Slides(3).Shapes(5).Table _

    .Cell(1, 1).Shape.TextFrame.TextRange _

    .Text = "Cell 1"
```

Use the [AddTable](PowerPoint.Shapes.AddTable.md)method to add a table to a slide. This example adds a 3x3 table on slide two in the active presentation.




```vb
ActivePresentation.Slides(2).Shapes.AddTable(3, 3)
```


## Methods



|Name|
|:-----|
|[ApplyStyle](PowerPoint.Table.ApplyStyle.md)|
|[Cell](PowerPoint.Table.Cell.md)|
|[ScaleProportionally](PowerPoint.Table.ScaleProportionally.md)|

## Properties



|Name|
|:-----|
|[AlternativeText](PowerPoint.Table.AlternativeText.md)|
|[Application](PowerPoint.Table.Application.md)|
|[Background](PowerPoint.Table.Background.md)|
|[Columns](PowerPoint.Table.Columns.md)|
|[FirstCol](PowerPoint.Table.FirstCol.md)|
|[FirstRow](PowerPoint.Table.FirstRow.md)|
|[HorizBanding](PowerPoint.Table.HorizBanding.md)|
|[LastCol](PowerPoint.Table.LastCol.md)|
|[LastRow](PowerPoint.Table.LastRow.md)|
|[Parent](PowerPoint.Table.Parent.md)|
|[Rows](PowerPoint.Table.Rows.md)|
|[Style](PowerPoint.Table.Style.md)|
|[TableDirection](PowerPoint.Table.TableDirection.md)|
|[Title](PowerPoint.Table.Title.md)|
|[VertBanding](PowerPoint.Table.VertBanding.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
