---
title: Error object (Excel)
keywords: vbaxl10.chm701072
f1_keywords:
- vbaxl10.chm701072
ms.prod: excel
api_name:
- Excel.Error
ms.assetid: bc8c4e3c-c831-58fd-c367-4246ad510ba9
ms.date: 03/29/2019
localization_priority: Normal
---

# Error object (Excel)

Represents a spreadsheet error for a range.


## Remarks

This object works for ranges containing only one cell.

Use the **[Item](Excel.Errors.Item.md)** property of the **Errors** object to return an **Error** object.

After an **Error** object is returned, you can use the **Value** property in conjunction with the **[Errors](Excel.Range.Errors.md)** property of the **Range** object to check whether a particular error checking option is enabled.

> [!NOTE] 
> Be careful not to confuse the **Error** object with the error handling features of Visual Basic.


## Example

The following example creates a formula in cell A1 referencing empty cells, and then it uses **Item** (_index_), where _index_ identifies the error type, to display a message stating the situation.

```vb
Sub CheckEmptyCells() 
 
 Dim rngFormula As Range 
 Set rngFormula = Application.Range("A1") 
 
 ' Place a formula referencing empty cells. 
 Range("A1").Formula = "=A2+A3" 
 Application.ErrorCheckingOptions.EmptyCellReferences = True 
 
 ' Perform check to see if EmptyCellReferences check is on. 
 If rngFormula.Errors.Item(xlEmptyCellReferences).Value = True Then 
 MsgBox "The empty cell references error checking feature is enabled." 
 Else 
 MsgBox "The empty cell references error checking feature is not on." 
 End If 
 
End Sub
```


## Properties

- [Application](Excel.Error.Application.md)
- [Creator](Excel.Error.Creator.md)
- [Ignore](Excel.Error.Ignore.md)
- [Parent](Excel.Error.Parent.md)
- [Value](Excel.Error.Value.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
