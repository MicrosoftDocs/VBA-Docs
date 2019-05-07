---
title: PivotFilter.IsMemberPropertyFilter property (Excel)
keywords: vbaxl10.chm770085
f1_keywords:
- vbaxl10.chm770085
ms.prod: excel
api_name:
- Excel.PivotFilter.IsMemberPropertyFilter
ms.assetid: 94b8055f-c45b-90fe-fd65-418f29e78ff0
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotFilter.IsMemberPropertyFilter property (Excel)

Specifies whether the label filter is based on the PivotItem captions of a member property of the field or on the PivotItem captions of the PivotField itself. Read-only **Boolean**.


## Syntax

_expression_.**IsMemberPropertyFilter**

_expression_ A variable that represents a **[PivotFilter](Excel.PivotFilter.md)** object.


## Remarks

The default value of this property is **False**.

Returns **True** if the label filter is based on PivotItem captions of a member property of the PivotField, or returns **False** if the filter is based on the PivotItem captions of the PivotField. This property is valid only for Label filters and only for OLAP PivotFields that have at least one member property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]