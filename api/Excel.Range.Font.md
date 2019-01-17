---
title: Range.Font property (Excel)
keywords: vbaxl10.chm144131
f1_keywords:
- vbaxl10.chm144131
ms.prod: excel
api_name:
- Excel.Range.Font
ms.assetid: d9cb8667-6c71-d311-a6e5-1d30d5718050
ms.date: 06/08/2017
localization_priority: Priority
---


# Range.Font property (Excel)

Returns a  **[Font](Excel.Font(object).md)** object that represents the font of the specified object.


## Syntax

_expression_. `Font`

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Example

This example determines the if the font name for cell A1 is Arial and notifies the user.


```vb
Sub CheckFont() 
 
 Range("A1").Select 
 
 ' Determine if the font name for selected cell is Arial. 
 If Range("A1").Font.Name = "Arial" Then 
 MsgBox "The font name for this cell is 'Arial'" 
 Else 
 MsgBox "The font name for this cell is not 'Arial'" 
 End If 
 
End Sub
```


## See also


[Range Object](Excel.Range(object).md)

