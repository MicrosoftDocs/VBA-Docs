---
title: Application.ThousandsSeparator property (Excel)
keywords: vbaxl10.chm133289
f1_keywords:
- vbaxl10.chm133289
ms.prod: excel
api_name:
- Excel.Application.ThousandsSeparator
ms.assetid: da244add-1c85-4636-2aff-b26feec215f3
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.ThousandsSeparator property (Excel)

Sets or returns the character used for the thousands separator as a **String**. Read/write.


## Syntax

_expression_.**ThousandsSeparator**

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