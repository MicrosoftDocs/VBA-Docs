---
title: Form.PivotTableChange property (Access)
keywords: vbaac10.chm13538,vbaac10.chm5102
f1_keywords:
- vbaac10.chm13538,vbaac10.chm5102
ms.prod: access
api_name:
- Access.Form.PivotTableChange
ms.assetid: d8d6a7eb-2bc1-e441-95fe-aefaec7fde9d
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.PivotTableChange property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[PivotTableChange](access.form.pivottablechange(even).md)** event occurs. Read/write.


## Syntax

_expression_.**PivotTableChange**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **PivotTableChange** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **PivotTableChange** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).PivotTableChange = "[Event Procedure]" 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]