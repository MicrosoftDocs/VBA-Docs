---
title: Form.BeforeScreenTip property (Access)
keywords: vbaac10.chm13547,vbaac10.chm5112
f1_keywords:
- vbaac10.chm13547,vbaac10.chm5112
ms.prod: access
api_name:
- Access.Form.BeforeScreenTip
ms.assetid: 4829b972-de4e-f8dc-f19c-c6a52c7dd14b
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.BeforeScreenTip property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[BeforeScreenTip](Access.Form.BeforeScreenTip(even).md)** event occurs. Read/write.


## Syntax

_expression_.**BeforeScreenTip**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **BeforeScreenTip** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **BeforeScreenTip** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).BeforeScreenTip = "[Event Procedure]"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]