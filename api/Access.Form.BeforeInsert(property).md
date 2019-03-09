---
title: Form.BeforeInsert property (Access)
keywords: vbaac10.chm13432
f1_keywords:
- vbaac10.chm13432
ms.prod: access
api_name:
- Access.Form.BeforeInsert
ms.assetid: 634b0480-ddb3-7ef7-b347-57ca9a4eebad
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.BeforeInsert property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[BeforeInsert](Access.Form.BeforeInsert(even).md)** event occurs. Read/write.


## Syntax

_expression_.**BeforeInsert**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **BeforeInsert** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **BeforeInsert** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).BeforeInsert = "[Event Procedure]" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]