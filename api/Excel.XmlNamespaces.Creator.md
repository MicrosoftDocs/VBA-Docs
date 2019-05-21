---
title: XmlNamespaces.Creator property (Excel)
keywords: vbaxl10.chm745074
f1_keywords:
- vbaxl10.chm745074
ms.prod: excel
api_name:
- Excel.XmlNamespaces.Creator
ms.assetid: 3ce50d2b-2910-d6c7-96ea-fd664b3d5acc
ms.date: 05/21/2019
localization_priority: Normal
---


# XmlNamespaces.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[XmlNamespaces](Excel.XmlNamespaces.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]