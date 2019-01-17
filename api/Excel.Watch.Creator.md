---
title: Watch.Creator property (Excel)
keywords: vbaxl10.chm689074
f1_keywords:
- vbaxl10.chm689074
ms.prod: excel
api_name:
- Excel.Watch.Creator
ms.assetid: 32ceb2af-a620-3a2e-cc27-92165eb81d8f
ms.date: 06/08/2017
localization_priority: Normal
---


# Watch.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [Watch](./Excel.Watch.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[Watch Object](Excel.Watch.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]