---
title: DocumentProperty.Value property (Office)
keywords: vbaof11.chm250006
f1_keywords:
- vbaof11.chm250006
ms.prod: office
api_name:
- Office.DocumentProperty.Value
ms.assetid: 2d66f8f7-0dfd-e3df-168f-1ca0dfbb0e70
ms.date: 01/08/2019
localization_priority: Normal
---


# DocumentProperty.Value property (Office)

Gets or sets the value of a document property. Read/write.


## Syntax

_expression_.**Value**

_expression_ Required. A variable that represents a **[DocumentProperty](Office.DocumentProperty.md)** object.


## Remarks

If the container application doesn't define a value for one of the built-in document properties, reading the **Value** property for that document property causes an error.


## See also

- [DocumentProperty object members](overview/library-reference/documentproperty-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]