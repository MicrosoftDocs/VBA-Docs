---
title: Rows object (Word)
ms.prod: word
ms.assetid: cd83d0ef-f743-1886-54de-497017c5f542
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows object (Word)

A collection of  **[Row](Word.Row.md)** objects that represent the table rows in the specified selection, range, or table.


## Remarks

Use the **Rows** property to return the **Rows** collection. The following example centers rows in the first table in the active document between the left and right margins.


```vb
ActiveDocument.Tables(1).Rows.Alignment = wdAlignRowCenter
```

Use the **Add** method to add a row to a table. The following example inserts a row before the first row in the selection.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Rows.Add BeforeRow:=Selection.Rows(1) 
End If
```

Use  **Rows** (Index), where Index is the index number, to return a single **Row** object. The index number represents the position of the row in the selection, range, or table. The following example deletes the first row in the first table in the active document.




```vb
ActiveDocument.Tables(1).Rows(1).Delete
```


## Methods



|Name|
|:-----|
|[Add](Word.Rows.Add.md)|
|[ConvertToText](Word.Rows.ConvertToText.md)|
|[Delete](Word.Rows.Delete.md)|
|[DistributeHeight](Word.Rows.DistributeHeight.md)|
|[Item](Word.Rows.Item.md)|
|[Select](Word.Rows.Select.md)|
|[SetHeight](Word.Rows.SetHeight.md)|
|[SetLeftIndent](Word.Rows.SetLeftIndent.md)|

## Properties



|Name|
|:-----|
|[Alignment](Word.Rows.Alignment.md)|
|[AllowBreakAcrossPages](Word.Rows.AllowBreakAcrossPages.md)|
|[AllowOverlap](Word.Rows.AllowOverlap.md)|
|[Application](Word.Rows.Application.md)|
|[Borders](Word.Rows.Borders.md)|
|[Count](Word.Rows.Count.md)|
|[Creator](Word.Rows.Creator.md)|
|[DistanceBottom](Word.Rows.DistanceBottom.md)|
|[DistanceLeft](Word.Rows.DistanceLeft.md)|
|[DistanceRight](Word.Rows.DistanceRight.md)|
|[DistanceTop](Word.Rows.DistanceTop.md)|
|[First](Word.Rows.First.md)|
|[HeadingFormat](Word.Rows.HeadingFormat.md)|
|[Height](Word.Rows.Height.md)|
|[HeightRule](Word.Rows.HeightRule.md)|
|[HorizontalPosition](Word.Rows.HorizontalPosition.md)|
|[Last](Word.Rows.Last.md)|
|[LeftIndent](Word.Rows.LeftIndent.md)|
|[NestingLevel](Word.Rows.NestingLevel.md)|
|[Parent](Word.Rows.Parent.md)|
|[RelativeHorizontalPosition](Word.Rows.RelativeHorizontalPosition.md)|
|[RelativeVerticalPosition](Word.Rows.RelativeVerticalPosition.md)|
|[Shading](Word.Rows.Shading.md)|
|[SpaceBetweenColumns](Word.Rows.SpaceBetweenColumns.md)|
|[TableDirection](Word.Rows.TableDirection.md)|
|[VerticalPosition](Word.Rows.VerticalPosition.md)|
|[WrapAroundText](Word.Rows.WrapAroundText.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
