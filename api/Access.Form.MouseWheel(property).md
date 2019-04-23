---
title: Form.MouseWheel property (Access)
keywords: vbaac10.chm13552
f1_keywords:
- vbaac10.chm13552
ms.prod: access
api_name:
- Access.Form.MouseWheel
ms.assetid: 364f7854-d7d5-5fe2-effa-6154e86376b4
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.MouseWheel property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[MouseWheel](access.form.mousewheel(even).md)** event occurs. Read/write.


## Syntax

_expression_.**MouseWheel**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **MouseWheel** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **MouseWheel** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).MouseWheel = "[Event Procedure]" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]