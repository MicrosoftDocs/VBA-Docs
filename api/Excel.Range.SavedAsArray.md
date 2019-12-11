---
title: Range.SavedAsArray property (Excel)
keywords: ????
f1_keywords:
- ???
ms.prod: excel
api_name:
- Excel.Range.SavedAsArray
ms.assetid: ????
ms.date: 12/10/2019
localization_priority: Normal
---


# Range.SavedAsArray property (Excel)

**True** if all of the cells in the range would be saved to file as an array formula; **False** if none of the cells in the range would be saved to file as a legacy array formula; **null** otherwise. Read-only **Variant**.


## Syntax

_expression_.**SavedAsArray**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.

## Remarks

Dynamic array Excel may save formulas and their associated spilled cells to file as an array formula to ensure they will calculate correctly in Pre-Dynamic array versions of Excel. These cells will appear as Legacy Array formulas in Pre-dynamic array Excel. 

## Example

This example prompts the user to select a range on Sheet1. If every cell in the selected range contains a spill, the example displays a message.

```vb
Worksheets("Sheet1").Activate 
Set rr = Application.InputBox( _ 
 prompt:="Select a range on this worksheet", _ 
 Type:=8) 
If rr.SavedAsArray = True Then 
 MsgBox "Every cell in the selection is part of a spilled range" 
End If
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
