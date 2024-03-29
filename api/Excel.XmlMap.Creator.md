---
title: XmlMap.Creator property (Excel)
keywords: vbaxl10.chm753074
f1_keywords:
- vbaxl10.chm753074
api_name:
- Excel.XmlMap.Creator
ms.assetid: a66d485c-8d92-edee-63dc-13c70d5faa53
ms.date: 05/21/2019
ms.localizationpriority: medium
---


# XmlMap.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[XmlMap](Excel.XmlMap.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]