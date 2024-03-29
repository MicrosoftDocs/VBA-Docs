---
title: Charts.Creator property (Excel)
keywords: vbaxl10.chm216074
f1_keywords:
- vbaxl10.chm216074
api_name:
- Excel.Charts.Creator
ms.assetid: 520db104-5cf3-c130-4590-e92b6b5e0d3e
ms.date: 04/20/2019
ms.localizationpriority: medium
---


# Charts.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[Charts](Excel.Charts.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]