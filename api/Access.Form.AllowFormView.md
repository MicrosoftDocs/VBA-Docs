---
title: Form.AllowFormView property (Access)
keywords: vbaac10.chm13532
f1_keywords:
- vbaac10.chm13532
ms.prod: access
api_name:
- Access.Form.AllowFormView
ms.assetid: 15dc69fc-d4ba-c8e3-d047-71f96c32fe02
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.AllowFormView property (Access)

Returns or sets a **Boolean** indicating whether the specified form may be viewed in Form view. **True** if Form view is allowed. Read/write.


## Syntax

_expression_.**AllowFormView**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Use the **[AllowDatasheetView](Access.Form.AllowDatasheetView.md)**, **AllowFormView**, **[AllowPivotChartView](Access.Form.AllowPivotChartView.md)**, or **[AllowPivotTableView](Access.Form.AllowPivotTableView.md)** properties to control which views are allowed for a form.


## Example

The following example makes Form view valid for the specified form, and then opens the form in Form view.

```vb
Forms(0).AllowFormView = True 
DoCmd.OpenForm FormName:=Forms(0).Name, View:=acNormal
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]