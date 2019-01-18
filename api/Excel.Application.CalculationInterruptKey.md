---
title: Application.CalculationInterruptKey property (Excel)
keywords: vbaxl10.chm133266
f1_keywords:
- vbaxl10.chm133266
ms.prod: excel
api_name:
- Excel.Application.CalculationInterruptKey
ms.assetid: 1187c122-0498-a82c-5479-1595c7f06448
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CalculationInterruptKey property (Excel)

Sets or returns an  **[xlCalculationInterruptKey](Excel.XlCalculationInterruptKey.md)** constant that specifies the key that can interrupt Microsoft Excel when performing calculations. Read/write.


## Syntax

_expression_. `CalculationInterruptKey`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

In this example, Microsoft Excel determines the setting for the calculation interrupt key and notifies the user.


```vb
Sub CheckInterruptKey() 
 
 ' Determine the calculation interrupt key and notify the user. 
 Select Case Application.CalculationInterruptKey 
 Case xlAnyKey 
 MsgBox "The calcuation interrupt key is set to any key." 
 Case xlEscKey 
 MsgBox "The calcuation interrupt key is set to 'Escape'" 
 Case xlNoKey 
 MsgBox "The calcuation interrupt key is set to no key." 
 End Select 
 
End Sub
```


## See also


[Application Object](Excel.Application(object).md)

