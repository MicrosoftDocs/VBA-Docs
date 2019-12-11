---
title: Range.HasSpill property (Excel)
keywords: ????
f1_keywords:
- ???
ms.prod: excel
api_name:
- Excel.Range.HasSpill
ms.assetid: ????
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.HasSpill property (Excel)

**True** if all of the cells in the range are part of a spilled range; **False** if none of the cells in the range are part of a spilled range; **null** otherwise. Read-only **Variant**.


## Syntax

_expression_.**HasSpill**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example prompts the user to select a range on Sheet1. If every cell in the selected range contains a spill, the example displays a message.

```vb
Worksheets("Sheet1").Activate 
Set rr = Application.InputBox( _ 
 prompt:="Select a range on this worksheet", _ 
 Type:=8) 
If rr.HasSpill = True Then 
 MsgBox "Every cell in the selection is part of a spilled range" 
End If
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
