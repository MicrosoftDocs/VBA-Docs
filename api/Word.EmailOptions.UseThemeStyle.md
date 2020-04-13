---
title: EmailOptions.UseThemeStyle property (Word)
keywords: vbawd10.chm165347431
f1_keywords:
- vbawd10.chm165347431
ms.prod: word
api_name:
- Word.EmailOptions.UseThemeStyle
ms.assetid: e34f27c6-4222-aa9a-dfbc-40c7c5c55a67
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.UseThemeStyle property (Word)

 **True** if new email messages use the character style defined by the default email message theme. Read/write **Boolean**.


## Syntax

_expression_. `UseThemeStyle`

_expression_ A variable that represents a '[EmailOptions](Word.EmailOptions.md)' object.


## Remarks

If no default email message theme has been specified, the **UseThemeStyle** property has no effect.


## Example

This example sets Microsoft Word to use the Artsy theme as the default theme for new email messages and to use the character style defined in the Artsy theme.


```vb
Application.EmailOptions.ThemeName = "artsy" 
Application.EmailOptions.UseThemeStyle = True
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]