---
title: Form.OnConnect property (Access)
keywords: vbaac10.chm13536,vbaac10.chm5100
f1_keywords:
- vbaac10.chm13536,vbaac10.chm5100
ms.prod: access
api_name:
- Access.Form.OnConnect
ms.assetid: de181e49-ccba-52fa-f521-3e55f3ed78d2
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.OnConnect property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[OnConnect](Access.Form.OnConnect(even).md)** event occurs. Read/write.


## Syntax

_expression_.**OnConnect**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **OnConnect** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **OnConnect** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).OnConnect = "[Event Procedure]" 

```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]