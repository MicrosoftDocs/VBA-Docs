---
title: SparkHorizontalAxis.Creator property (Excel)
keywords: vbaxl10.chm878074
f1_keywords:
- vbaxl10.chm878074
ms.prod: excel
api_name:
- Excel.SparkHorizontalAxis.Creator
ms.assetid: 2420bb92-b37e-cbd2-22d3-906cdc039afa
ms.date: 06/08/2017
localization_priority: Normal
---


# SparkHorizontalAxis.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a '[SparkHorizontalAxis](Excel.SparkHorizontalAxis.md)' object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[SparkHorizontalAxis Object](Excel.SparkHorizontalAxis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]