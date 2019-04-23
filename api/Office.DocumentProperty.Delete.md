---
title: DocumentProperty.Delete method (Office)
keywords: vbaof11.chm250004
f1_keywords:
- vbaof11.chm250004
ms.prod: office
api_name:
- Office.DocumentProperty.Delete
ms.assetid: 2a9ac097-0156-007f-2b4b-62a34b240f71
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentProperty.Delete method (Office)

Removes a custom document property.


## Syntax

_expression_.**Delete**

_expression_ Required. A variable that represents a **[DocumentProperty](Office.DocumentProperty.md)** object.


## Remarks

You cannot delete a built-in document property.


## Example

This example deletes a custom document property. For this example to run properly, you must have a custom **DocumentProperty** object named **CustomNumber**.

```vb
ActiveDocument.CustomDocumentProperties("CustomNumber").Delete
```


## See also

- [DocumentProperty object members](overview/library-reference/documentproperty-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]