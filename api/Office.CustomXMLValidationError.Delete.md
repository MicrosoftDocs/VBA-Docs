---
title: CustomXMLValidationError.Delete method (Office)
keywords: vbaof11.chm307006
f1_keywords:
- vbaof11.chm307006
ms.prod: office
api_name:
- Office.CustomXMLValidationError.Delete
ms.assetid: d425c0f8-6eb1-9e1d-5246-3ba77bbf3cd3
ms.date: 01/07/2019
localization_priority: Normal
---


# CustomXMLValidationError.Delete method (Office)

Deletes the **CustomXMLValidationError** object representing a data validation error.


## Syntax

_expression_.**Delete**

_expression_ An expression that returns a **[CustomXMLValidationError](Office.CustomXMLValidationError.md)** object.


## Example

The following example deletes the validation error containing specific text.


```vb
Dim objCustomXMLValidationError as CustomXMLValidationError 
 
' Deletes the specified error message. 
objCustomXMLValidationError.Text("To add content to this data stream, it must be valid, well-formed XML.").Delete
```


## See also

- [CustomXMLValidationError object members](overview/library-reference/customxmlvalidationerror-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]