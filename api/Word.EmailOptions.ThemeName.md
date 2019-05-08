---
title: EmailOptions.ThemeName property (Word)
keywords: vbawd10.chm165347442
f1_keywords:
- vbawd10.chm165347442
ms.prod: word
api_name:
- Word.EmailOptions.ThemeName
ms.assetid: ec988c2a-9cf3-867c-81f4-cfa6d00b54d9
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.ThemeName property (Word)

Returns or sets the name of the theme plus any theme formatting options to use for new email messages. Read/write  **String**.


## Syntax

_expression_. `ThemeName`

_expression_ A variable that represents a '[EmailOptions](Word.EmailOptions.md)' object.


## Remarks

For an explanation of the value returned by this property, see the Name argument of the  **[ApplyTheme](Word.Document.ApplyTheme.md)** method. The value returned by this property may not correspond to the theme's display name as it appears in the Theme dialog box. To return a theme's display name, use the **[ActiveThemeDisplayName](Word.Document.ActiveThemeDisplayName.md)** property.

You can also use the  **[GetDefaultTheme](Word.Application.GetDefaultTheme.md)** and **[SetDefaultTheme](Word.Application.SetDefaultTheme.md)** methods to return and set the default theme for new email messages.


## Example

This example sets Microsoft Word to use the Blueprint theme with Vivid Colors for all new email messages.


```vb
Application.EmailOptions.ThemeName = "blueprnt 100"
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]