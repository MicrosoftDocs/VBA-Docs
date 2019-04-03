---
title: ConditionalStyle object (Word)
keywords: vbawd10.chm1389
f1_keywords:
- vbawd10.chm1389
ms.prod: word
api_name:
- Word.ConditionalStyle
ms.assetid: 2380494e-09e9-8494-a93c-8bbaf621aad1
ms.date: 06/08/2017
localization_priority: Normal
---


# ConditionalStyle object (Word)

Represents special formatting applied to specified areas of a table when the selected table is formatted with a specified table style.


## Remarks

Use the  **[Condition](Word.TableStyle.Condition.md)** method of the **[TableStyle](Word.TableStyle.md)** object to return a **ConditionalStyle** object. The **Shading** property can be used to apply shading to specified areas of a table. This example selects the first table in the active document and applies shading to alternate rows and columns. This example assumes that there is a table in the active document and that it is formatted using the Table Grid style.


```vb
Sub ApplyConditionalStyle() 
 With ActiveDocument 
 .Tables(1).Select 
 With .Styles("Table Grid").Table 
 .Condition(wdOddColumnBanding).Shading _ 
 .BackgroundPatternColor = wdColorGray10 
 .Condition(wdOddRowBanding).Shading _ 
 .BackgroundPatternColor = wdColorGray10 
 End With 
 End With 
End Sub
```

Use the  **[Borders](Word.TableStyle.Borders.md)** property to apply borders to specified areas of a table. This example selects the first table in the active document and applies borders to the first and last row and first column. This example assumes that there is a table in the active document and that it is formatted using the Table Grid style.




```vb
Sub ApplyTableBorders() 
 With ActiveDocument 
 .Tables(1).Select 
 With .Styles("Table Grid").Table 
 .Condition(wdFirstRow).Borders(wdBorderBottom) _ 
 .LineStyle = wdLineStyleDouble 
 .Condition(wdFirstColumn).Borders(wdBorderRight) _ 
 .LineStyle = wdLineStyleDouble 
 .Condition(wdLastRow).Borders(wdBorderTop) _ 
 .LineStyle = wdLineStyleDouble 
 End With 
 End With 
End Sub
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]