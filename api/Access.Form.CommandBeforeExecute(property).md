---
title: Form.CommandBeforeExecute property (Access)
keywords: vbaac10.chm13542
f1_keywords:
- vbaac10.chm13542
ms.prod: access
api_name:
- Access.Form.CommandBeforeExecute
ms.assetid: 574568fa-e488-6d4d-a42f-07eb7c7f9536
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.CommandBeforeExecute property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[CommandBeforeExecute](Access.Form.CommandBeforeExecute(even).md)** event occurs. Read/write.


## Syntax

_expression_.**CommandBeforeExecute**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **CommandBeforeExecute** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **CommandBeforeExecute** event occurs on the first form of the current project, the associated event procedure should run.

```vb
Forms(0).CommandBeforeExecute = "[Event Procedure]"
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]