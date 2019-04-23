---
title: Form.BeforeRender property (Access)
keywords: vbaac10.chm13551
f1_keywords:
- vbaac10.chm13551
ms.prod: access
api_name:
- Access.Form.BeforeRender
ms.assetid: f80035ac-4ce6-ac8a-203f-c36afab5cd01
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.BeforeRender property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[BeforeRender](Access.Form.BeforeRender(even).md)** event occurs. Read/write.


## Syntax

_expression_.**BeforeRender**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **BeforeRender** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **BeforeRender** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).BeforeRender = "[Event Procedure]"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]