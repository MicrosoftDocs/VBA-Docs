---
title: SlicerCaches.Creator property (Excel)
keywords: vbaxl10.chm894074
f1_keywords:
- vbaxl10.chm894074
ms.prod: excel
api_name:
- Excel.SlicerCaches.Creator
ms.assetid: e7e2e448-189a-051d-33f2-0dbb8de272d5
ms.date: 06/08/2017
localization_priority: Normal
---


# SlicerCaches.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a '[SlicerCaches](Excel.SlicerCaches.md)' collection.


## Return value

 **Long**


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[SlicerCaches Object](Excel.SlicerCaches.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]