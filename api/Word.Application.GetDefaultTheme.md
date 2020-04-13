---
title: Application.GetDefaultTheme method (Word)
keywords: vbawd10.chm158335392
f1_keywords:
- vbawd10.chm158335392
ms.prod: word
api_name:
- Word.Application.GetDefaultTheme
ms.assetid: 967760c0-4f99-5fae-026d-5ac60358d21c
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GetDefaultTheme method (Word)

Returns a  **String** that represents the name of the default theme plus the theme formatting options Microsoft Word uses for new documents, email messages, or Web pages.


## Syntax

_expression_. `GetDefaultTheme`( `_DocumentType_` )

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DocumentType_|Required| **WdDocumentMedium**|The type of new document for which you want to retrieve the default theme name.|

## Remarks

You can also use the **ThemeName** property to return and set the default theme for new email messages.


## Example

This example displays the name of the theme Word uses for new Web pages.


```vb
MsgBox Application.GetDefaultTheme(wdWebPage)
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]