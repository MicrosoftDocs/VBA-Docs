---
title: Validation.ErrorTitle property (Excel)
keywords: vbaxl10.chm532080
f1_keywords:
- vbaxl10.chm532080
ms.prod: excel
api_name:
- Excel.Validation.ErrorTitle
ms.assetid: bafa328c-9f2f-3bb3-be61-5772e28fed47
ms.date: 06/08/2017
localization_priority: Normal
---


# Validation.ErrorTitle property (Excel)

Returns or sets the title of the data-validation error dialog box. Read/write  **String**.


## Syntax

_expression_. `ErrorTitle`

_expression_ A variable that represents a [Validation](./Excel.Validation.md) object.


## Example

This example adds data validation to cell E5.


```vb
With Range("e5").Validation 
 .Add xlValidateWholeNumber, _ 
 xlValidAlertInformation, xlBetween, "5", "10" 
 .InputTitle = "Integers" 
 .ErrorTitle = "Integers" 
 .InputMessage = "Enter an integer from five to ten" 
 .ErrorMessage = "You must enter a number from five to ten" 
End With
```


## See also


[Validation Object](Excel.Validation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]