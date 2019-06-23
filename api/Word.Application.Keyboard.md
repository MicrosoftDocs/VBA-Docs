---
title: Application.Keyboard method (Word)
keywords: vbawd10.chm158335422
f1_keywords:
- vbawd10.chm158335422
ms.prod: word
api_name:
- Word.Application.Keyboard
ms.assetid: 67745d17-3dec-b4d9-919e-49925f2a7e34
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Keyboard method (Word)

Returns or sets the keyboard language and layout settings.


## Syntax

_expression_. `Keyboard`( `_LangId_` )

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _LangId_|Optional| **Long**|The language and layout combination to which Microsoft Word sets the keyboard. If this argument is omitted, the method returns the current language and layout setting.|

## Remarks

Microsoft Windows tracks keyboard language and layout settings using a variable type called an input language handle, often referred to as an HKL. The low word of the handle is a language ID, and the high word is a handle to a keyboard layout.


## Example

This example assigns the current keyboard language and layout setting to a variable.


```vb
Dim lngKeyboard As Long 
 
lng
```


```vb
Keyboard = Application.Keyboard
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]