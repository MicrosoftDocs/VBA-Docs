---
title: Form.DataChange property (Access)
keywords: vbaac10.chm13554
f1_keywords:
- vbaac10.chm13554
ms.prod: access
api_name:
- Access.Form.DataChange
ms.assetid: 14fd4c9c-eb18-8f4d-ebd9-6f389523c4cf
ms.date: 03/09/2019
localization_priority: Normal
---


# Form.DataChange property (Access)

Returns or sets a **String** indicating which macro, event procedure, or user-defined function runs when the **[DataChange](Access.Form.DataChange(even).md)** event occurs. Read/write.


## Syntax

_expression_.**DataChange**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values for this property are:

- _macroname_, where _macroname_ is the name of a macro.

- [Event Procedure], which indicates the event procedure associated with the **DataChange** event for the specified object.

- _=functionname()_, where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the **DataChange** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).DataChange = "[Event Procedure]" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]