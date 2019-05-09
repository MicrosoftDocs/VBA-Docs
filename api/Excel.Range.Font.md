---
title: Range.Font property (Excel)
keywords: vbaxl10.chm144131
f1_keywords:
- vbaxl10.chm144131
ms.prod: excel
api_name:
- Excel.Range.Font
ms.assetid: d9cb8667-6c71-d311-a6e5-1d30d5718050
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.Font property (Excel)

Returns a **[Font](Excel.Font(object).md)** object that represents the font of the specified object.


## Syntax

_expression_.**Font**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example determines if the font name for cell A1 is Arial and notifies the user.

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




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
