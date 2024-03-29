---
title: ChartCategory.Creator property (Excel)
keywords: vbaxl10.chm945074
f1_keywords:
- vbaxl10.chm945074
ms.assetid: 2e9f59f5-bfd2-9518-f34a-705216b85c3f
ms.date: 04/20/2019
ms.localizationpriority: medium
---


# ChartCategory.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ChartCategory](Excel.chartcategory.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## Property value

**XLCREATOR**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]