---
title: Form.CommandChecked property (Access)
keywords: vbaac10.chm13543
f1_keywords:
- vbaac10.chm13543
ms.prod: access
api_name:
- Access.Form.CommandChecked
ms.assetid: 4f3bb0fa-6f3f-4836-a0d0-06d480e1d194
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.CommandChecked property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[CommandChecked](Access.Form.CommandChecked(even).md)** event occurs. Read/write.


## Syntax

_expression_.**CommandChecked**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **CommandChecked** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **CommandChecked** event occurs on the first form of the current project, the associated event procedure should run.

```vb
Forms(0).CommandChecked = "[Event Procedure]" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]