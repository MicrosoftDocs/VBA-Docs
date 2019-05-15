---
title: SlicerItems.Creator property (Excel)
keywords: vbaxl10.chm908074
f1_keywords:
- vbaxl10.chm908074
ms.prod: excel
api_name:
- Excel.SlicerItems.Creator
ms.assetid: d7002e14-3c07-3255-6b01-556fc1d3c503
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerItems.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SlicerItems](Excel.SlicerItems.md)** object.


## Return value

**Long**


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[SlicerItems Object](Excel.SlicerItems.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]