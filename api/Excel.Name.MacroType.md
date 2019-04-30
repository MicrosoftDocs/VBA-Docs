---
title: Name.MacroType property (Excel)
keywords: vbaxl10.chm490078
f1_keywords:
- vbaxl10.chm490078
ms.prod: excel
api_name:
- Excel.Name.MacroType
ms.assetid: 46f02cb6-56c3-7b0e-27a4-db356802abe6
ms.date: 05/01/2019
localization_priority: Normal
---


# Name.MacroType property (Excel)

Returns or sets what the name refers to. Read/write **[XlXLMMacroType](Excel.XlXLMMacroType.md)**.


## Syntax

_expression_.**MacroType**

_expression_ A variable that represents a **[Name](Excel.Name.md)** object.


## Example

This example assumes that you created a custom function or command on a Microsoft Excel version 4.0 macro sheet. The example displays the function category, in the language of the macro. It assumes that the name of the custom function or command is the only name in the workbook.

```vb
With ActiveWorkbook.Names(1) 
 If .MacroType <> xlNotXLM Then 
 MsgBox "The category for this name is " & .Category 
 Else 
 MsgBox "This name does not refer to" & _ 
 " a custom function or command." 
 End If 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]