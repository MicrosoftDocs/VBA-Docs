---
title: SparkVerticalAxis.Creator property (Excel)
keywords: vbaxl10.chm880074
f1_keywords:
- vbaxl10.chm880074
api_name:
- Excel.SparkVerticalAxis.Creator
ms.assetid: 931a6fd8-57cb-ca6f-44a6-aff2d5a2dfcb
ms.date: 05/16/2019
ms.localizationpriority: medium
---


# SparkVerticalAxis.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SparkVerticalAxis](Excel.SparkVerticalAxis.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]