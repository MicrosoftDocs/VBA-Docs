---
title: CheckBox.AfterUpdate property (Access)
keywords: vbaac10.chm10736
f1_keywords:
- vbaac10.chm10736
ms.prod: access
api_name:
- Access.CheckBox.AfterUpdate
ms.assetid: eaef525d-4447-86b5-9567-311e7324b720
ms.date: 02/12/2019
localization_priority: Normal
---


# CheckBox.AfterUpdate property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the **[AfterUpdate](access.checkbox.afterupdate-event.md)** event occurs. Read/write **String**.


## Syntax

_expression_.**AfterUpdate**

_expression_ A variable that represents a **[CheckBox](Access.CheckBox.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **AfterUpdate** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **AfterUpdate** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).AfterUpdate = "[Event Procedure]" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]