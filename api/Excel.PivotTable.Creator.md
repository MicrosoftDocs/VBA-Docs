---
title: PivotTable.Creator property (Excel)
keywords: vbaxl10.chm234074
f1_keywords:
- vbaxl10.chm234074
api_name:
- Excel.PivotTable.Creator
ms.assetid: 7066bafd-10d6-f4f3-4236-40bd942a1c39
ms.date: 05/08/2019
ms.localizationpriority: medium
---


# PivotTable.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]