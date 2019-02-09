---
title: Form.AfterUpdate property (Access)
keywords: vbaac10.chm13435
f1_keywords:
- vbaac10.chm13435
ms.prod: access
api_name:
- Access.Form.AfterUpdate
ms.assetid: 5002727c-24bc-4067-0e5e-3c63b8b6427e
ms.date: 06/08/2017
localization_priority: Normal
---


# Form.AfterUpdate property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the  **AfterUpdate** event occurs. Read/write **String**.


## Syntax

_expression_.**AfterUpdate**

_expression_ A variable that represents a [Form](Access.Form.md) object.


## Remarks

Valid values for this property are: 

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **AfterUpdate** event for the specified object.

- **=** _functionname_ **()**, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the AfterUpdate event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).AfterUpdate = "[Event Procedure]" 

```


## See also


[Form Object](Access.Form.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]