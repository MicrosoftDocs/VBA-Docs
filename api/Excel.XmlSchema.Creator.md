---
title: XmlSchema.Creator property (Excel)
keywords: vbaxl10.chm749074
f1_keywords:
- vbaxl10.chm749074
api_name:
- Excel.XmlSchema.Creator
ms.assetid: d255b385-bc2f-84ca-68f3-79fe2c250651
ms.date: 05/21/2019
ms.localizationpriority: medium
---


# XmlSchema.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[XmlSchema](Excel.XmlSchema.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]