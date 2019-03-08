---
title: Form.ViewChange property (Access)
keywords: vbaac10.chm13553,vbaac10.chm5118
f1_keywords:
- vbaac10.chm13553,vbaac10.chm5118
ms.prod: access
api_name:
- Access.Form.ViewChange
ms.assetid: f8a8fe82-6983-5632-b779-879faf228ac2
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.ViewChange property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[ViewChange](Access.Form.ViewChange(even).md)** event occurs. Read/write.


## Syntax

_expression_.**ViewChange**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **ViewChange** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **ViewChange** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).ViewChange = "[Event Procedure]" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]