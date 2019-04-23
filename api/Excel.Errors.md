---
title: Errors object (Excel)
keywords: vbaxl10.chm699072
f1_keywords:
- vbaxl10.chm699072
ms.prod: excel
api_name:
- Excel.Errors
ms.assetid: d2b50bbf-2685-fc5f-74c5-fa8bb9955f2a
ms.date: 03/29/2019
localization_priority: Normal
---


# Errors object (Excel)

Represents the various spreadsheet errors for a range.


## Remarks

Use the **[Errors](Excel.Range.Errors.md)** property of the **Range** object to return an **Errors** object.


## Example

After an **Errors** object is returned, you can use the **Value** property of the **[Error](Excel.Error.md)** object to check for particular error-checking conditions. The following example places a number as text in cell A1, and then notifies the user when the value of cell A1 contains a number as text.

```vb
Sub ErrorValue() 
 
 ' Place a number written as text in cell A1. 
 Range("A1").Formula = "'1" 
 
 If Range("A1").Errors.Item(xlNumberAsText).Value = True Then 
 MsgBox "Cell A1 has a number as text." 
 Else 
 MsgBox "Cell A1 is a number." 
 End If 
 
End Sub
```

## Properties

- [Application](Excel.Errors.Application.md)
- [Creator](Excel.Errors.Creator.md)
- [Item](Excel.Errors.Item.md)
- [Parent](Excel.Errors.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]