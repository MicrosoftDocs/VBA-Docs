---
title: ModelMeasureName.Creator property (Excel)
keywords: vbaxl10.chm969074
f1_keywords:
- vbaxl10.chm969074
ms.prod: excel
ms.assetid: 60c5ed37-0a61-76e8-fc5e-2c5fdf084cbe
ms.date: 05/01/2019
localization_priority: Normal
---


# ModelMeasureName.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ModelMeasureName](Excel.modelmeasurename.md)** object.


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string XCEL.


## Property value

**XLCREATOR**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]