---
title: XmlNamespace.Creator property (Excel)
keywords: vbaxl10.chm743074
f1_keywords:
- vbaxl10.chm743074
api_name:
- Excel.XmlNamespace.Creator
ms.assetid: ceed703f-148b-607f-f6cf-13484aeff5da
ms.date: 05/21/2019
ms.localizationpriority: medium
---


# XmlNamespace.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[XmlNamespace](Excel.XmlNamespace.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]