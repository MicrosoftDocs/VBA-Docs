---
title: Table Object (Word)
keywords: vbawd10.chm2385
f1_keywords:
- vbawd10.chm2385
ms.prod: word
api_name:
- Word.Table
ms.assetid: 996b58dd-ebc6-ee30-5bfe-c5e51a0f71d6
ms.date: 06/08/2017
---


# Table Object (Word)

Represents a single table. The  **Table** object is a member of the **[Tables](../../../api/Word.tables.md)** collection. The **Tables** collection includes all the tables in the specified selection, range, or document.


## Remarks

Use  **Tables** (Index), where Index is the index number, to return a single **Table** object. The index number represents the position of the table in the selection, range, or document. The following example converts the first table in the active document to text.


```
ActiveDocument.Tables(1).ConvertToText Separator:=wdSeparateByTabs
```

Use the  **Add** method to add a table at the specified range. The following example adds a 3x4 table at the beginning of the active document.




```
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.Tables.Add Range:=myRange, NumRows:=3, NumColumns:=4
```


## Methods



|**Name**|
|:-----|
|[ApplyStyleDirectFormatting](../../../api/Word.Table.ApplyStyleDirectFormatting.md)|
|[AutoFitBehavior](../../../api/Word.Table.AutoFitBehavior.md)|
|[AutoFormat](../../../api/Word.Table.AutoFormat.md)|
|[Cell](../../../api/Word.Table.Cell.md)|
|[ConvertToText](../../../api/Word.Table.ConvertToText.md)|
|[Delete](../../../api/Word.Table.Delete.md)|
|[Select](../../../api/Word.Table.Select.md)|
|[Sort](../../../api/Word.Table.Sort.md)|
|[SortAscending](../../../api/Word.Table.SortAscending.md)|
|[SortDescending](../../../api/Word.Table.SortDescending.md)|
|[Split](../../../api/Word.Table.Split.md)|
|[UpdateAutoFormat](../../../api/Word.Table.UpdateAutoFormat.md)|

## Properties



|**Name**|
|:-----|
|[AllowAutoFit](../../../api/Word.Table.AllowAutoFit.md)|
|[Application](../../../api/Word.Table.Application.md)|
|[ApplyStyleColumnBands](../../../api/Word.Table.ApplyStyleColumnBands.md)|
|[ApplyStyleFirstColumn](../../../api/Word.Table.ApplyStyleFirstColumn.md)|
|[ApplyStyleHeadingRows](../../../api/Word.Table.ApplyStyleHeadingRows.md)|
|[ApplyStyleLastColumn](../../../api/Word.Table.ApplyStyleLastColumn.md)|
|[ApplyStyleLastRow](../../../api/Word.Table.ApplyStyleLastRow.md)|
|[ApplyStyleRowBands](../../../api/Word.Table.ApplyStyleRowBands.md)|
|[AutoFormatType](../../../api/Word.Table.AutoFormatType.md)|
|[Borders](../../../api/Word.Table.Borders.md)|
|[BottomPadding](../../../api/Word.Table.BottomPadding.md)|
|[Columns](../../../api/Word.Table.Columns.md)|
|[Creator](../../../api/Word.Table.Creator.md)|
|[Descr](../../../api/Word.Table.Descr.md)|
|[ID](../../../api/Word.Table.ID.md)|
|[LeftPadding](../../../api/Word.Table.LeftPadding.md)|
|[NestingLevel](../../../api/Word.Table.NestingLevel.md)|
|[Parent](../../../api/Word.Table.Parent.md)|
|[PreferredWidth](../../../api/Word.Table.PreferredWidth.md)|
|[PreferredWidthType](../../../api/Word.Table.PreferredWidthType.md)|
|[Range](../../../api/Word.Table.Range.md)|
|[RightPadding](../../../api/Word.Table.RightPadding.md)|
|[Rows](../../../api/Word.Table.Rows.md)|
|[Shading](../../../api/Word.Table.Shading.md)|
|[Spacing](../../../api/Word.Table.Spacing.md)|
|[Style](../../../api/Word.Table.Style.md)|
|[TableDirection](../../../api/Word.Table.TableDirection.md)|
|[Tables](../../../api/Word.Table.Tables.md)|
|[Title](../../../api/Word.Table.Title.md)|
|[TopPadding](../../../api/Word.Table.TopPadding.md)|
|[Uniform](../../../api/Word.Table.Uniform.md)|

## See also


#### Other resources


[Word Object Model Reference](../../../api/overview/object-model-word-vba-reference.md)

