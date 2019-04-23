---
title: Form.BeforeQuery property (Access)
keywords: vbaac10.chm13540
f1_keywords:
- vbaac10.chm13540
ms.prod: access
api_name:
- Access.Form.BeforeQuery
ms.assetid: 40e763fd-897a-a0b1-72a9-d73ec628e397
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.BeforeQuery property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[BeforeQuery](Access.Form.BeforeQuery(even).md)** event occurs. Read/write.


## Syntax

_expression_.**BeforeQuery**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **BeforeQuery** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **BeforeQuery** event occurs on the first form of the current project, the associated event procedure should run.

```vb
Forms(0).BeforeQuery = "[Event Procedure]" 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]