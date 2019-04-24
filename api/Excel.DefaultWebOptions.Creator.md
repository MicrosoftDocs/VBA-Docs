---
title: DefaultWebOptions.Creator property (Excel)
keywords: vbaxl10.chm659074
f1_keywords:
- vbaxl10.chm659074
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.Creator
ms.assetid: 5fcbd08f-1e37-db2c-75c2-db65c4af3f00
ms.date: 04/25/2019
localization_priority: Normal
---


# DefaultWebOptions.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[DefaultWebOptions](Excel.DefaultWebOptions.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]