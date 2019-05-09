---
title: PublishObject.Creator property (Excel)
keywords: vbaxl10.chm651074
f1_keywords:
- vbaxl10.chm651074
ms.prod: excel
api_name:
- Excel.PublishObject.Creator
ms.assetid: 9f579e1f-3943-e116-bbe4-3ef58dc9179e
ms.date: 05/09/2019
localization_priority: Normal
---


# PublishObject.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[PublishObject](Excel.PublishObject.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]