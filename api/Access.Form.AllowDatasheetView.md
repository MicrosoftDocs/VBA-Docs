---
title: Form.AllowDatasheetView property (Access)
keywords: vbaac10.chm13533
f1_keywords:
- vbaac10.chm13533
ms.prod: access
api_name:
- Access.Form.AllowDatasheetView
ms.assetid: 81796b90-94dd-cd27-3613-a2050e2bce21
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.AllowDatasheetView property (Access)

Returns or sets a **Boolean** indicating whether the specified form may be viewed in Datasheet view. **True** if Datasheet view is allowed. Read/write.


## Syntax

_expression_.**AllowDatasheetView**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Use the **AllowDatasheetView**, **[AllowFormView](Access.Form.AllowFormView.md)**, **[AllowPivotChartView](Access.Form.AllowPivotChartView.md)**, or **[AllowPivotTableView](Access.Form.AllowPivotTableView.md)** properties to control which views are allowed for a form.


## Example

The following example makes Datasheet view valid for the specified form, and then opens the form in Datasheet view.

```vb
Forms(0).AllowDatasheetView = True 
DoCmd.OpenForm FormName:=Forms(0).Name, View:=acFormDS 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]