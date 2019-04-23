---
title: Form.CommandExecute property (Access)
keywords: vbaac10.chm13545
f1_keywords:
- vbaac10.chm13545
ms.prod: access
api_name:
- Access.Form.CommandExecute
ms.assetid: b105b107-8123-5cfe-b87d-cb53518e3dba
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.CommandExecute property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[CommandExecute](Access.Form.CommandExecute(even).md)** event occurs. Read/write.


## Syntax

_expression_.**CommandExecute**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **CommandExecute** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **CommandExecute** event occurs on the first form of the current project, the associated event procedure should run.

```vb
Forms(0).CommandExecute = "[Event Procedure]"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]