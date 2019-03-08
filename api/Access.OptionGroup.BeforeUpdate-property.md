---
title: OptionGroup.BeforeUpdate property (Access)
keywords: vbaac10.chm10862
f1_keywords:
- vbaac10.chm10862
ms.prod: access
api_name:
- Access.OptionGroup.BeforeUpdate
ms.assetid: 0ea86e13-03ba-9f56-ef42-e8147fa70064
ms.date: 02/12/2019
localization_priority: Normal
---


# OptionGroup.BeforeUpdate property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the **[BeforeUpdate](access.OptionGroup.beforeupdate-event.md)** event occurs. Read/write **String**.


## Syntax

_expression_.**BeforeUpdate**

_expression_ A variable that represents an **[OptionGroup](Access.OptionGroup.md)** object.


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