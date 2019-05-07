---
title: PivotField.UseMemberPropertyAsCaption property (Excel)
keywords: vbaxl10.chm240139
f1_keywords:
- vbaxl10.chm240139
ms.prod: excel
api_name:
- Excel.PivotField.UseMemberPropertyAsCaption
ms.assetid: 4e5e9a53-c746-25db-78c0-115282851829
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotField.UseMemberPropertyAsCaption property (Excel)

This property is used to control whether member property captions are used for PivotItem captions of the PivotField. Read/write **Boolean**.


## Syntax

_expression_.**UseMemberPropertyAsCaption**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

If **UseMemberPropertyAsCaption** is set to **True** for a PivotField, **[MemberPropertyCaption](excel.pivotfield.memberpropertycaption.md)** specifies which member property caption to display. If none is specified, the first member property of that PivotField (in data source order) will be displayed as the caption for items of that PivotField.

If **UseMemberPropertyAsCaption** is set to **False**, the regular PivotItem captions are used for the PivotField.

If you try to set **UseMemberPropertyAsCaption** to **True** for a PivotField with no member properties, a run-time error is returned. For PivotFields with no member properties, the property is always **False**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]