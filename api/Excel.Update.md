---
title: Update method (Excel Graph)
keywords: vbagr10.chm66216
f1_keywords:
- vbagr10.chm66216
ms.prod: excel
api_name:
- Excel.Update
ms.assetid: ef26d691-e77a-115e-2152-eec136aa6839
ms.date: 04/09/2019
localization_priority: Normal
---


# Update method (Excel Graph)

Updates the specified embedded object in the host file.

## Syntax

_expression_.**Update**

_expression_ Required. An expression that returns an **[Application](Excel.Application-graph-object.md)** object.


## Example

This example updates the application.

```vb
Sub UseUpdate() 
 
 Application.Update 
 
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]