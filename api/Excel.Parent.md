---
title: Parent property (Excel Graph)
keywords: vbagr10.chm65686
f1_keywords:
- vbagr10.chm65686
ms.prod: excel
api_name:
- Excel.Parent
ms.assetid: 504783e9-8bd6-7716-20d4-1f1484f36b33
ms.date: 04/11/2019
localization_priority: Normal
---


# Parent property (Excel Graph)

Returns the parent object.

## Syntax

_expression_.**Parent**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example returns the parent object of the application.

```vb
Sub UseParent() 
 
 Application.Parent 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]