---
title: Application.SubstituteFont method (Word)
keywords: vbawd10.chm158335280
f1_keywords:
- vbawd10.chm158335280
ms.prod: word
api_name:
- Word.Application.SubstituteFont
ms.assetid: 2563bf9a-31ea-4104-b26b-538eb7e27f85
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SubstituteFont method (Word)

Sets font-mapping options.


## Syntax

_expression_. `SubstituteFont`( `_UnavailableFont_` , `_SubstituteFont_` )

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _UnavailableFont_|Required| **String**|The name of a font not available on your computer that you want to map to a different font for display and printing.|
| _SubstituteFont_|Required| **String**|The name of a font available on your computer that you want to substitute for the unavailable font.|

## Remarks

You can find font-mapping options in the  **Font Substitution** dialog box.


## Example

This example substitutes Courier for CustomFont1.


```vb
Application.SubstituteFont UnavailableFont:= "CustomFont1", _ 
 SubstituteFont:= "Courier"
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]