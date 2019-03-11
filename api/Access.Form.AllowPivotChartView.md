---
title: Form.AllowPivotChartView property (Access)
keywords: vbaac10.chm13535
f1_keywords:
- vbaac10.chm13535
ms.prod: access
api_name:
- Access.Form.AllowPivotChartView
ms.assetid: 5585b530-d114-d07e-63cb-8d96dec458e8
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.AllowPivotChartView property (Access)

Returns or sets a **Boolean** indicating whether the specified form may be viewed in PivotChart view. **True** if PivotChart view is allowed. Read/write.


## Syntax

_expression_.**AllowPivotChartView**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Use the **[AllowDatasheetView](Access.Form.AllowDatasheetView.md)**, **[AllowFormView](Access.Form.AllowFormView.md)**, **AllowPivotChartView**, or **[AllowPivotTableView](Access.Form.AllowPivotTableView.md)** properties to control which views are allowed for a form.


## Example

The following example makes PivotChart view valid for the specified form, and then opens the form in PivotChart view.

```vb
Forms(0).AllowPivotChartView = True 
DoCmd.OpenForm FormName:=Forms(0).Name, View:=acFormPivotChart 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]