---
title: Form.AfterFinalRender property (Access)
keywords: vbaac10.chm13548,vbaac10.chm5113
f1_keywords:
- vbaac10.chm13548,vbaac10.chm5113
ms.prod: access
api_name:
- Access.Form.AfterFinalRender
ms.assetid: c6e294f8-8cd9-1413-eff8-f2b033766326
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.AfterFinalRender property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[AfterFinalRender](Access.Form.AfterFinalRender(even).md)** event occurs. Read/write.


## Syntax

_expression_.**AfterFinalRender**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **AfterFinalRender** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **AfterFinalRender** event occurs on the first form of the current project, the associated event procedure should run.

```vb
Forms(0).AfterFinalRender = "[Event Procedure]" 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]