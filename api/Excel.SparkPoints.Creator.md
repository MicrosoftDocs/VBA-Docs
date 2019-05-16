---
title: SparkPoints.Creator property (Excel)
keywords: vbaxl10.chm872074
f1_keywords:
- vbaxl10.chm872074
ms.prod: excel
api_name:
- Excel.SparkPoints.Creator
ms.assetid: 65ad69c7-3c71-f844-2cef-325d707a225d
ms.date: 05/16/2019
localization_priority: Normal
---


# SparkPoints.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SparkPoints](Excel.SparkPoints.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]