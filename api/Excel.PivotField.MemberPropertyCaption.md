---
title: PivotField.MemberPropertyCaption property (Excel)
keywords: vbaxl10.chm240140
f1_keywords:
- vbaxl10.chm240140
ms.prod: excel
api_name:
- Excel.PivotField.MemberPropertyCaption
ms.assetid: 66f2ad5f-cd37-74ef-e9df-cd4793212026
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotField.MemberPropertyCaption property (Excel)

Setting the **MemberPropertyCaption** property controls which member property is used as a caption for a given level. Read/write **Boolean**.


## Syntax

_expression_.**MemberPropertyCaption**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

This setting has a visual effect only when **[UseMemberPropertyAsCaption](excel.pivotfield.usememberpropertyascaption.md)** is set to **True** for the PivotField.

When **MemberPropertyCaption** is set, the setting is remembered while toggling the **UseMemberPropertyAsCaption** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]