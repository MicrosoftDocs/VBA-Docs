---
title: Document.ApplyTheme method (Word)
keywords: vbawd10.chm158007618
f1_keywords:
- vbawd10.chm158007618
ms.prod: word
api_name:
- Word.Document.ApplyTheme
ms.assetid: a4b9180e-5128-6a19-a629-47c20837f84b
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ApplyTheme method (Word)

Applies a theme to an open document.


## Syntax

_expression_. `ApplyTheme`( `_Name_` )

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the theme plus any theme formatting options you want to apply. The format of this string is "themennn" where theme and nnn are defined as follows:

|**String**|**Description**|
|:-----|:-----|
|theme|The name of the folder that contains the data for the requested theme. (The default location for theme data folders is C:\Program Files\Common Files\Microsoft Shared\Themes.) You must use the folder name for the theme rather than the display name that appears in the **Theme** dialog box.|
|nnn|A three-digit string that indicates which theme formatting options to activate (1 to activate, 0 to deactivate). The digits correspond to the **Vivid Colors**,  **Active Graphics**, and  **Background Image** check boxes in the **Theme** dialog box. If this string is omitted, the default value for nnn is "011" (Active Graphics and Background Image are activated).|
|

## Example

This example applies the Artsy theme to the active document and activates the Vivid Colors option.


```vb
ActiveDocument.ApplyTheme "artsy 100"
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]