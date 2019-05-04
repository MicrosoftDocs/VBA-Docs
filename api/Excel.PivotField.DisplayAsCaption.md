---
title: PivotField.DisplayAsCaption property (Excel)
keywords: vbaxl10.chm240143
f1_keywords:
- vbaxl10.chm240143
ms.prod: excel
api_name:
- Excel.PivotField.DisplayAsCaption
ms.assetid: b2eadf78-2b5b-69cf-7929-fba28608de38
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.DisplayAsCaption property (Excel)

This property is used to display member properties of PivotFields as captions. Read-only.


## Syntax

_expression_.**DisplayAsCaption**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

This property returns **True** when a given member property is used as a caption, and **False** when a member property PivotField is not used as a caption. Trying to use this property for PivotFields that are not member properties will return a run-time error.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]