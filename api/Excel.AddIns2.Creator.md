---
title: AddIns2.Creator property (Excel)
keywords: vbaxl10.chm866074
f1_keywords:
- vbaxl10.chm866074
ms.prod: excel
api_name:
- Excel.AddIns2.Creator
ms.assetid: bd20266f-a3d8-58da-505b-48f905896fb6
ms.date: 04/03/2019
localization_priority: Normal
---


# AddIns2.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[AddIns2](Excel.AddIns2.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]