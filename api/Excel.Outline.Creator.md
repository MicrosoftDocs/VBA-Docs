---
title: Outline.Creator property (Excel)
keywords: vbaxl10.chm454074
f1_keywords:
- vbaxl10.chm454074
api_name:
- Excel.Outline.Creator
ms.assetid: b0d9637e-c913-54c1-f782-7f933e4b39dd
ms.date: 05/02/2019
ms.localizationpriority: medium
---


# Outline.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[Outline](Excel.Outline.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]