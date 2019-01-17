---
title: ConnectorFormat.Creator property (Excel)
keywords: vbaxl10.chm645074
f1_keywords:
- vbaxl10.chm645074
ms.prod: excel
api_name:
- Excel.ConnectorFormat.Creator
ms.assetid: ba6891ca-344f-25d9-1430-a32652fed7b3
ms.date: 06/08/2017
localization_priority: Normal
---


# ConnectorFormat.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [ConnectorFormat](Excel.ConnectorFormat.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[ConnectorFormat Object](Excel.ConnectorFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]