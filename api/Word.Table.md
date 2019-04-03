---
title: Table object (Word)
keywords: vbawd10.chm2385
f1_keywords:
- vbawd10.chm2385
ms.prod: word
api_name:
- Word.Table
ms.assetid: 996b58dd-ebc6-ee30-5bfe-c5e51a0f71d6
ms.date: 06/08/2017
localization_priority: Normal
---


# Table object (Word)

Represents a single table. The  **Table** object is a member of the **[Tables](./Word.tables.md)** collection. The **Tables** collection includes all the tables in the specified selection, range, or document.


## Remarks

Use  **Tables** (Index), where Index is the index number, to return a single **Table** object. The index number represents the position of the table in the selection, range, or document. The following example converts the first table in the active document to text.


```vb
ActiveDocument.Tables(1).ConvertToText Separator:=wdSeparateByTabs
```

Use the  **Add** method to add a table at the specified range. The following example adds a 3x4 table at the beginning of the active document.




```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.Tables.Add Range:=myRange, NumRows:=3, NumColumns:=4
```


## Methods



|Name|
|:-----|
|[ApplyStyleDirectFormatting](./Word.Table.ApplyStyleDirectFormatting.md)|
|[AutoFitBehavior](./Word.Table.AutoFitBehavior.md)|
|[AutoFormat](./Word.Table.AutoFormat.md)|
|[Cell](./Word.Table.Cell.md)|
|[ConvertToText](./Word.Table.ConvertToText.md)|
|[Delete](./Word.Table.Delete.md)|
|[Select](./Word.Table.Select.md)|
|[Sort](./Word.Table.Sort.md)|
|[SortAscending](./Word.Table.SortAscending.md)|
|[SortDescending](./Word.Table.SortDescending.md)|
|[Split](./Word.Table.Split.md)|
|[UpdateAutoFormat](./Word.Table.UpdateAutoFormat.md)|

## Properties



|Name|
|:-----|
|[AllowAutoFit](./Word.Table.AllowAutoFit.md)|
|[Application](./Word.Table.Application.md)|
|[ApplyStyleColumnBands](./Word.Table.ApplyStyleColumnBands.md)|
|[ApplyStyleFirstColumn](./Word.Table.ApplyStyleFirstColumn.md)|
|[ApplyStyleHeadingRows](./Word.Table.ApplyStyleHeadingRows.md)|
|[ApplyStyleLastColumn](./Word.Table.ApplyStyleLastColumn.md)|
|[ApplyStyleLastRow](./Word.Table.ApplyStyleLastRow.md)|
|[ApplyStyleRowBands](./Word.Table.ApplyStyleRowBands.md)|
|[AutoFormatType](./Word.Table.AutoFormatType.md)|
|[Borders](./Word.Table.Borders.md)|
|[BottomPadding](./Word.Table.BottomPadding.md)|
|[Columns](./Word.Table.Columns.md)|
|[Creator](./Word.Table.Creator.md)|
|[Descr](./Word.Table.Descr.md)|
|[ID](./Word.Table.ID.md)|
|[LeftPadding](./Word.Table.LeftPadding.md)|
|[NestingLevel](./Word.Table.NestingLevel.md)|
|[Parent](./Word.Table.Parent.md)|
|[PreferredWidth](./Word.Table.PreferredWidth.md)|
|[PreferredWidthType](./Word.Table.PreferredWidthType.md)|
|[Range](./Word.Table.Range.md)|
|[RightPadding](./Word.Table.RightPadding.md)|
|[Rows](./Word.Table.Rows.md)|
|[Shading](./Word.Table.Shading.md)|
|[Spacing](./Word.Table.Spacing.md)|
|[Style](./Word.Table.Style.md)|
|[TableDirection](./Word.Table.TableDirection.md)|
|[Tables](./Word.Table.Tables.md)|
|[Title](./Word.Table.Title.md)|
|[TopPadding](./Word.Table.TopPadding.md)|
|[Uniform](./Word.Table.Uniform.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
