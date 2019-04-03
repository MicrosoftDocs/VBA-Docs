---
title: Application.DecimalSeparator property (Excel)
keywords: vbaxl10.chm133288
f1_keywords:
- vbaxl10.chm133288
ms.prod: excel
api_name:
- Excel.Application.DecimalSeparator
ms.assetid: 2423d0dd-2b67-e8d2-c611-2bd3c8061f66
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DecimalSeparator property (Excel)

Sets or returns the character used for the decimal separator as a **String**. Read/write.


## Syntax

_expression_.**DecimalSeparator**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example places "1,234,567.89" in cell A1, and then changes the system separators to dashes for the decimals and thousands separators.

```vb
Sub ChangeSystemSeparators() 
 
 Range("A1").Formula = "1,234,567.89" 
 MsgBox "The system separators will now change." 
 
 ' Define separators and apply. 
 Application.DecimalSeparator = "-" 
 Application.ThousandsSeparator = "-" 
 Application.UseSystemSeparators = False 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
