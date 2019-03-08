---
title: BoundObjectFrame.BeforeUpdate property (Access)
keywords: vbaac10.chm10961
f1_keywords:
- vbaac10.chm10961
ms.prod: access
api_name:
- Access.BoundObjectFrame.BeforeUpdate
ms.assetid: 01ee3c67-76c6-b651-042b-a7aa59e7443e
ms.date: 02/08/2019
localization_priority: Normal
---


# BoundObjectFrame.BeforeUpdate property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the **[BeforeUpdate](access.boundobjectframe.beforeupdate-event.md)** event occurs. Read/write **String**.


## Syntax

_expression_.**BeforeUpdate**

_expression_ A variable that represents a **[BoundObjectFrame](Access.BoundObjectFrame.md)** object.


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