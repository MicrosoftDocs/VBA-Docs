---
title: TableStyle object (Word)
keywords: vbawd10.chm3735
f1_keywords:
- vbawd10.chm3735
ms.prod: word
api_name:
- Word.TableStyle
ms.assetid: 4f1f4489-0ef7-dff0-8f2a-77f87937f3ad
ms.date: 06/08/2017
localization_priority: Normal
---


# TableStyle object (Word)

Represents a single style that can be applied to a table.


## Remarks

Use the  **Table** property of the **Styles** object to return a **TableStyle** object. Use the **Borders** property to apply borders to an entire table. Use the **Condition** method to apply borders or shading only to specified sections of a table. This example creates a new table style and formats the table with a surrounding border. Special borders and shading are applied to the first and last rows and the last column.


```vb
Sub NewTableStyle() 
 Dim styTable As Style 
 
 Set styTable = ActiveDocument.Styles.Add( _ 
 Name:="TableStyle 1", Type:=wdStyleTypeTable) 
 
 With styTable.Table 
 
 'Apply borders around table 
 .Borders(wdBorderTop).LineStyle = wdLineStyleSingle 
 .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle 
 .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle 
 .Borders(wdBorderRight).LineStyle = wdLineStyleSingle 
 
 'Apply a double border to the heading row 
 .Condition(wdFirstRow).Borders(wdBorderBottom) _ 
 .LineStyle = wdLineStyleDouble 
 
 'Apply a double border to the last column 
 .Condition(wdLastColumn).Borders(wdBorderLeft) _ 
 .LineStyle = wdLineStyleDouble 
 
 'Apply shading to last row 
 .Condition(wdLastRow).Shading _ 
 .BackgroundPatternColor = wdColorGray125 
 
 End With 
 
End Sub
```


## Methods



|Name|
|:-----|
|[Condition](Word.TableStyle.Condition.md)|

## Properties



|Name|
|:-----|
|[Alignment](Word.TableStyle.Alignment.md)|
|[AllowBreakAcrossPage](Word.TableStyle.AllowBreakAcrossPage.md)|
|[AllowPageBreaks](Word.TableStyle.AllowPageBreaks.md)|
|[Application](Word.TableStyle.Application.md)|
|[Borders](Word.TableStyle.Borders.md)|
|[BottomPadding](Word.TableStyle.BottomPadding.md)|
|[ColumnStripe](Word.TableStyle.ColumnStripe.md)|
|[Creator](Word.TableStyle.Creator.md)|
|[LeftIndent](Word.TableStyle.LeftIndent.md)|
|[LeftPadding](Word.TableStyle.LeftPadding.md)|
|[Parent](Word.TableStyle.Parent.md)|
|[RightPadding](Word.TableStyle.RightPadding.md)|
|[RowStripe](Word.TableStyle.RowStripe.md)|
|[Shading](Word.TableStyle.Shading.md)|
|[Spacing](Word.TableStyle.Spacing.md)|
|[TableDirection](Word.TableStyle.TableDirection.md)|
|[TopPadding](Word.TableStyle.TopPadding.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]