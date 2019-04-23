---
title: Index property (Excel Graph)
keywords: vbagr10.chm66022
f1_keywords:
- vbagr10.chm66022
ms.prod: excel
api_name:
- Excel.Index
ms.assetid: 39e1b38c-776c-fd78-0115-a14672d022f2
ms.date: 04/11/2019
localization_priority: Normal
---


# Index property (Excel Graph)

Returns the index number of the object within the collection of similar objects. Read-only **Long**.

## Syntax

_expression_.**Index**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example displays the index number of an object passed to this procedure.

```vb
MsgBox "The index number of this object is " & obj.Index
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]