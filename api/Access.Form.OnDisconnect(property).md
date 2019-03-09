---
title: Form.OnDisconnect property (Access)
keywords: vbaac10.chm13537,vbaac10.chm5101
f1_keywords:
- vbaac10.chm13537,vbaac10.chm5101
ms.prod: access
api_name:
- Access.Form.OnDisconnect
ms.assetid: 8f6514c7-8f61-2ae7-0859-8299523609ca
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.OnDisconnect property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[OnDisconnect](Access.Form.OnDisconnect(even).md)** event occurs. Read/write.


## Syntax

_expression_.**OnDisconnect**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **OnDisconnect** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **OnDisconnect** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).OnDisconnect = "[Event Procedure]" 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]