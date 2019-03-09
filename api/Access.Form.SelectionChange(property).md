---
title: Form.SelectionChange property (Access)
keywords: vbaac10.chm13541
f1_keywords:
- vbaac10.chm13541
ms.prod: access
api_name:
- Access.Form.SelectionChange
ms.assetid: e31876fc-103a-d231-a6fa-7cb026a343e1
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.SelectionChange property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[SelectionChange](Access.Form.SelectionChange(even).md)** event occurs. Read/write.


## Syntax

_expression_.**SelectionChange**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **SelectionChange** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **SelectionChange** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).SelectionChange = "[Event Procedure]" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]