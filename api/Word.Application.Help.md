---
title: Application.Help method (Word)
keywords: vbawd10.chm158335305
f1_keywords:
- vbawd10.chm158335305
ms.prod: word
api_name:
- Word.Application.Help
ms.assetid: ff64e6bd-e29b-7cfc-437b-df8b8e59ce59
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Help method (Word)

Displays installed Help information.


## Syntax

_expression_. `Help`( `_HelpType_` )

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _HelpType_|Required| **Variant**|The on-line Help topic or window. Can be any of these  **[WdHelpType](Word.WdHelpType.md)** constants: **wdHelp**, **wdHelpAbout**, **wdHelpActiveWindow**, **wdHelpContents**, **wdHelpHWP**, **wdHelpIchitaro**, **wdHelpIndex**, **wdHelpPE2**, **wdHelpPSSHelp**, **wdHelpSearch**, **wdHelpUsingHelp**. (Some of the constants listed here may not be available to you, depending on the language that you have selected or installed.)|

## Example

This example displays the **Help Topics** dialog box.


```vb
Help HelpType:=wdHelp
```

This example displays a list of Help topics that describe how to use Help.




```vb
Help HelpType:=wdHelpUsingHelp
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]