---
title: Form.AfterInsert property (Access)
keywords: vbaac10.chm13433
f1_keywords:
- vbaac10.chm13433
ms.prod: access
api_name:
- Access.Form.AfterInsert
ms.assetid: 95bc1f0d-a0fa-ffdd-ef5a-e6eb2a854feb
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.AfterInsert property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[AfterInsert](Access.Form.AfterInsert(even).md)** event occurs. Read/write.


## Syntax

_expression_.**AfterInsert**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **AfterInsert** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **AfterInsert** event occurs on the first form of the current project, the associated event procedure should run.

```vb
Forms(0).AfterInsert = "[Event Procedure]" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]