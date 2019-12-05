---
title: Range.SpillParent property (Excel)
keywords: ????
f1_keywords:
- ???
ms.prod: excel
api_name:
- Excel.Range.SpillParent
ms.assetid: ????
ms.date: ????
localization_priority: Normal
---


# Range.SpillParent property (Excel)

If a range contains a spill, returns the cell containing the formula responsible. Otherwise an error is returned. 

## Syntax

_expression_.**SpillParent**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.

## Remarks

This was property was introduced to support Dynamic array Excel. Range.HasSpill can be used to determine if a cell is a member of a spill range.

## Example

This example prompts the user to select a range on Sheet1. If the selected cell is a part of a spill range, the originating cell is returned. Otherwise, the user is notified that the cell is not a part of a spill range.

```vb
Set rr = Application.InputBox( _
 prompt:="Select a cell on this worksheet", _
 Type:=8)
If rr.HasSpill = True Then
 MsgBox "The spill is coming from " & rr.SpillParent.Address
Else
 MsgBox "This cell is not part of a spill range"
End If
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
