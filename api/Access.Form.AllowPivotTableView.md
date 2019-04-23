---
title: Form.AllowPivotTableView property (Access)
keywords: vbaac10.chm13534,vbaac10.chm5540
f1_keywords:
- vbaac10.chm13534,vbaac10.chm5540
ms.prod: access
api_name:
- Access.Form.AllowPivotTableView
ms.assetid: 42bad4b4-7de1-f144-9482-2e114fc5cc4b
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.AllowPivotTableView property (Access)

Returns or sets a **Boolean** indicating whether the specified form may be viewed in PivotTable view. **True** if PivotTable view is allowed. Read/write.


## Syntax

_expression_.**AllowPivotTableView**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Use the **[AllowDatasheetView](Access.Form.AllowDatasheetView.md)**, **[AllowFormView](Access.Form.AllowFormView.md)**, **[AllowPivotChartView](Access.Form.AllowPivotChartView.md)**, or **AllowPivotTableView** properties to control which views are allowed for a form.


## Example

The following example makes PivotTable view valid for the specified form, and then opens the form in PivotTable view.

```vb
Forms(0).AllowPivotTableView = True 
DoCmd.OpenForm FormName:=Forms(0).Name, View:=acFormPivotTable 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]