---
title: Error Object (Excel)
keywords: vbaxl10.chm701072
f1_keywords:
- vbaxl10.chm701072
ms.prod: excel
api_name:
- Excel.Error
ms.assetid: bc8c4e3c-c831-58fd-c367-4246ad510ba9
ms.date: 06/08/2017
---

# Error Object (Excel)

Represents a spreadsheet error for a range.


## Remarks
This object works for ranges containing only one cell.

Use the  **[Item](Excel.Errors.Item.md)** property of the **[Errors](Excel.Errors.md)** object to return an **Error** object.

Once an  **Error** object is returned, you can use the **[Value](Excel.Error.Value.md)** property, in conjunction with the **[Errors](Excel.Range.Errors.md)** property to check whether a particular error checking option is enabled.


 **Note**  Be careful not to confuse the  **Error** object with error handling features of Visual Basic.


## Example

The following example creates a formula in cell A1 referencing empty cells, and then it uses  **Item** ( _index_ ), where _index_ identifies the error type, to display a message stating the situation.


```
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



|**Name**|
|:-----|
|[Application](Excel.Error.Application.md)|
|[Creator](Excel.Error.Creator.md)|
|[Ignore](Excel.Error.Ignore.md)|
|[Parent](Excel.Error.Parent.md)|
|[Value](Excel.Error.Value.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

