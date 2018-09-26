---
title: ChartFormat.Creator Property (Excel)
keywords: vbaxl10.chm860074
f1_keywords:
- vbaxl10.chm860074
ms.prod: excel
api_name:
- Excel.ChartFormat.Creator
ms.assetid: 17992dc8-ef3c-2bac-2c52-8523c71424b9
ms.date: 06/08/2017
---


# ChartFormat.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_. `Creator`

 _expression_ A variable that represents a [ChartFormat](Excel.ChartFormat.md) object.


### Return value

XlCreator


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[ChartFormat Object](Excel.ChartFormat.md)

