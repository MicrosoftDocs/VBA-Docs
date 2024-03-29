---
title: ShapeRange.Creator property (Excel)
keywords: vbaxl10.chm639074
f1_keywords:
- vbaxl10.chm639074
api_name:
- Excel.ShapeRange.Creator
ms.assetid: 5ac1fcc9-ad5c-f25b-2e09-a8f3febcacef
ms.date: 05/14/2019
ms.localizationpriority: medium
---


# ShapeRange.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]