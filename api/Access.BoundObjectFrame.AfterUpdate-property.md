---
title: BoundObjectFrame.AfterUpdate property (Access)
keywords: vbaac10.chm10962
f1_keywords:
- vbaac10.chm10962
ms.prod: access
api_name:
- Access.BoundObjectFrame.AfterUpdate
ms.assetid: 3ebda4de-49c3-bfe7-8743-1c2c98caca58
ms.date: 02/08/2019
localization_priority: Normal
---


# BoundObjectFrame.AfterUpdate property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the **[AfterUpdate](access.boundobjectframe.afterupdate-event.md)** event occurs. Read/write **String**.


## Syntax

_expression_.**AfterUpdate**

_expression_ A variable that represents a **[BoundObjectFrame](Access.BoundObjectFrame.md)** object.


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