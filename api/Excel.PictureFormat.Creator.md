---
title: PictureFormat.Creator property (Excel)
ms.prod: excel
api_name:
- Excel.PictureFormat.Creator
ms.assetid: 4a2777a6-ed15-ed24-4553-1b96172ab57f
ms.date: 06/08/2017
localization_priority: Normal
---


# PictureFormat.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [PictureFormat](Excel.PictureFormat.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[PictureFormat Object](Excel.PictureFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]