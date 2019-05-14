---
title: ShapeNodes.Creator property (Excel)
ms.prod: excel
api_name:
- Excel.ShapeNodes.Creator
ms.assetid: 995d9596-a48b-4fd2-6682-45c453ed91ad
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeNodes.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ShapeNodes](Excel.ShapeNodes.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]