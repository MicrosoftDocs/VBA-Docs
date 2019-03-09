---
title: Form.AfterLayout property (Access)
keywords: vbaac10.chm13550
f1_keywords:
- vbaac10.chm13550
ms.prod: access
api_name:
- Access.Form.AfterLayout
ms.assetid: 8d548e7b-6d68-4631-2c59-f6b8d39cbb12
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.AfterLayout property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[AfterLayout](Access.Form.AfterLayout(even).md)** event occurs. Read/write.


## Syntax

_expression_.**AfterLayout**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **AfterLayout** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **AfterLayout** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).AfterLayout = "[Event Procedure]" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]