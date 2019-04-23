---
title: DataBarBorder.Creator property (Excel)
keywords: vbaxl10.chm884074
f1_keywords:
- vbaxl10.chm884074
ms.prod: excel
api_name:
- Excel.DataBarBorder.Creator
ms.assetid: 2d240406-f29b-6014-4cc0-06085c9573d8
ms.date: 04/23/2019
localization_priority: Normal
---


# DataBarBorder.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[DataBarBorder](Excel.DataBarBorder.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]