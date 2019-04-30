---
title: Name.Category property (Excel)
keywords: vbaxl10.chm490075
f1_keywords:
- vbaxl10.chm490075
ms.prod: excel
api_name:
- Excel.Name.Category
ms.assetid: 01892c7b-a42e-e4b3-6ddd-27ace1c51aae
ms.date: 05/01/2019
localization_priority: Normal
---


# Name.Category property (Excel)

Returns or sets the category for the specified name in the language of the macro. The name must refer to a custom function or command. Read/write **String**.


## Syntax

_expression_.**Category**

_expression_ A variable that represents a **[Name](Excel.Name.md)** object.


## Example

This example assumes that you created a custom function or command on a Microsoft Excel 4.0 macro sheet. The example displays the function category in the language of the macro. It assumes that the name of the custom function or command is the only name in the workbook.

```vb
With ActiveWorkbook.Names(1) 
 If .MacroType <> xlNone Then 
 MsgBox "The category for this name is " & .Category 
 Else 
 MsgBox "This name does not refer to" & _ 
 " a custom function or command." 
 End If 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]