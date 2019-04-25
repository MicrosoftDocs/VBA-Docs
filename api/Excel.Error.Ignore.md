---
title: Error.Ignore property (Excel)
keywords: vbaxl10.chm702074
f1_keywords:
- vbaxl10.chm702074
ms.prod: excel
api_name:
- Excel.Error.Ignore
ms.assetid: 2e1eea04-fa93-86ed-670a-23246dddfbfe
ms.date: 04/26/2019
localization_priority: Normal
---


# Error.Ignore property (Excel)

Allows the user to set or return the state of an error checking option for a range. **False** enables an error checking option for a range. **True** disables an error checking option for a range. Read/write **Boolean**.


## Syntax

_expression_.**Ignore**

_expression_ A variable that represents an **[Error](Excel.Error.md)** object.


## Remarks

Reference the **[ErrorCheckingOptions](Excel.ErrorCheckingOptions.md)** object to view a list of index values associated with error checking options.


## Example

This example disables the ignore flag in cell A1 for checking empty cell references.

```vb
Sub IgnoreChecking() 
 
 Range("A1").Select 
 
 ' Determine if empty cell references error checking is on, if not turn it on. 
 If Application.Range("A1").Errors(xlEmptyCellReferences).Ignore = True Then 
 Application.Range("A1").Errors(xlEmptyCellReferences).Ignore = False 
 MsgBox "Empty cell references error checking has been enabled for cell A1." 
 Else 
 MsgBox "Empty cell references error checking is already enabled for cell A1." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
