---
title: Application.ErrorCheckingOptions property (Excel)
keywords: vbaxl10.chm133280
f1_keywords:
- vbaxl10.chm133280
ms.prod: excel
api_name:
- Excel.Application.ErrorCheckingOptions
ms.assetid: 3821c6fd-e6c2-70cc-f546-70fdac6a6161
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.ErrorCheckingOptions property (Excel)

Returns an **[ErrorCheckingOptions](Excel.ErrorCheckingOptions.md)** object, which represents the error checking options for an application.


## Syntax

_expression_.**ErrorCheckingOptions**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

In this example, the **TextDate** property is used in conjunction with the **ErrorCheckingOptions** property. When the user selects a cell containing a two-digit year in the date, the **AutoCorrect Options** button appears.


```vb
Sub CheckTextDate() 
 
 ' Enable Microsoft Excel to identify dates written as text. 
 Application.ErrorCheckingOptions.TextDate = True 
 
 Range("A1").Formula = "'April 23, 00" 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]