---
title: Icon.Creator property (Excel)
keywords: vbaxl10.chm815074
f1_keywords:
- vbaxl10.chm815074
api_name:
- Excel.Icon.Creator
ms.assetid: 5cdaa402-b79a-735e-0413-b48fb037b510
ms.date: 04/27/2019
ms.localizationpriority: medium
---


# Icon.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[Icon](Excel.Icon.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]