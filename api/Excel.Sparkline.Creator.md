---
title: Sparkline.Creator property (Excel)
keywords: vbaxl10.chm874074
f1_keywords:
- vbaxl10.chm874074
ms.prod: excel
api_name:
- Excel.Sparkline.Creator
ms.assetid: 8353b55b-5494-4101-b5e1-78b0f2fdf152
ms.date: 06/08/2017
localization_priority: Normal
---


# Sparkline.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a '[Sparkline](Excel.Sparkline.md)' object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[Sparkline Object](Excel.Sparkline.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]