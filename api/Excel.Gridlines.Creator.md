---
title: Gridlines.Creator property (Excel)
keywords: vbaxl10.chm601074
f1_keywords:
- vbaxl10.chm601074
ms.prod: excel
api_name:
- Excel.Gridlines.Creator
ms.assetid: 095a985e-3823-a483-59d5-82afd93f5a5e
ms.date: 06/08/2017
localization_priority: Normal
---


# Gridlines.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [Gridlines](Excel.Gridlines-graph-object.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[Gridlines Object](Excel.Gridlines(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]