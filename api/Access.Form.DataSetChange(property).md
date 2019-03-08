---
title: Form.DataSetChange property (Access)
keywords: vbaac10.chm13546,vbaac10.chm5111
f1_keywords:
- vbaac10.chm13546,vbaac10.chm5111
ms.prod: access
api_name:
- Access.Form.DataSetChange
ms.assetid: 29f7f9a8-4dbd-9f69-7f4c-7f93add9f1b6
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.DataSetChange property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[DataSetChange](Access.Form.DataSetChange(even).md)** event occurs. Read/write.


## Syntax

_expression_.**DataSetChange**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **DataSetChange** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **DataSetChange** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).DataSetChange = "[Event Procedure]" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]