---
title: Form.CommandEnabled property (Access)
keywords: vbaac10.chm13544,vbaac10.chm5108
f1_keywords:
- vbaac10.chm13544,vbaac10.chm5108
ms.prod: access
api_name:
- Access.Form.CommandEnabled
ms.assetid: 07e6989d-9739-e023-32e4-95147eb4bba3
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.CommandEnabled property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[CommandEnabled](Access.Form.CommandEnabled(even).md)** event occurs. Read/write.


## Syntax

_expression_.**CommandEnabled**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **CommandEnabled** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **CommandEnabled** event occurs on the first form of the current project, the associated event procedure should run.

```vb
Forms(0).CommandEnabled = "[Event Procedure]" 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]