---
title: ComboBox.BeforeUpdate property (Access)
keywords: vbaac10.chm11447
f1_keywords:
- vbaac10.chm11447
ms.prod: access
api_name:
- Access.ComboBox.BeforeUpdate
ms.assetid: ce748fb1-4f8d-9e96-f77c-5dfc54dfee48
ms.date: 02/12/2019
localization_priority: Normal
---


# ComboBox.BeforeUpdate property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the **[BeforeUpdate](access.ComboBox.beforeupdate-event.md)** event occurs. Read/write **String**.


## Syntax

_expression_.**BeforeUpdate**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **BeforeUpdate** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.

    

## Example

The following example specifies that when the **BeforeUpdate** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).BeforeUpdate = "[Event Procedure]" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]