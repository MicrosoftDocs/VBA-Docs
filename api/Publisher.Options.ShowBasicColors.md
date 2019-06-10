---
title: Options.ShowBasicColors property (Publisher)
keywords: vbapb10.chm1048601
f1_keywords:
- vbapb10.chm1048601
ms.prod: publisher
api_name:
- Publisher.Options.ShowBasicColors
ms.assetid: d04504fa-5627-b66b-bd6e-30556155632c
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.ShowBasicColors property (Publisher)

Returns or sets a **Boolean** indicating whether Microsoft Publisher shows basic colors in the color palette; **True** to show basic colors in the palette. Read/write.


## Syntax

_expression_.**ShowBasicColors**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

Boolean


## Example

The following example sets Publisher to not show basic colors in the color palette.

```vb
Options.ShowBasicColors = False
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]