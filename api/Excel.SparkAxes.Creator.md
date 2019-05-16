---
title: SparkAxes.Creator property (Excel)
keywords: vbaxl10.chm876074
f1_keywords:
- vbaxl10.chm876074
ms.prod: excel
api_name:
- Excel.SparkAxes.Creator
ms.assetid: a5a60f1d-acb4-18f2-178a-4a71338ad606
ms.date: 05/16/2019
localization_priority: Normal
---


# SparkAxes.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SparkAxes](Excel.SparkAxes.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]