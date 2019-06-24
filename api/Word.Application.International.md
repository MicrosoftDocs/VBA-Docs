---
title: Application.International property (Word)
keywords: vbawd10.chm158335022
f1_keywords:
- vbawd10.chm158335022
ms.prod: word
api_name:
- Word.Application.International
ms.assetid: 907c2908-01a6-2a83-9968-98c21b699f4b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.International property (Word)

Returns information about the current country/region and international settings. Read-only  **Variant**.


## Syntax

_expression_. `International` (_Index_)

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **WdInternationalIndex**|The current country/region and/or international setting.|

## Example

This example displays the currency format in the status bar.


```vb
StatusBar = "Currency Format: " _ 
 & Application.International(wdCurrencyCode)
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]