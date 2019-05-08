---
title: Global.Help method (Word)
keywords: vbawd10.chm163119433
f1_keywords:
- vbawd10.chm163119433
ms.prod: word
api_name:
- Word.Global.Help
ms.assetid: cfae6e61-84bf-2462-39c5-569baec866ee
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.Help method (Word)

Displays on-line Help information.


## Syntax

_expression_. `Help`( `_HelpType_` )

_expression_ A variable that represents a '[Global](Word.Global.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _HelpType_|Required| **Variant**|The on-line Help topic or window. Can be any of these  **[WdHelpType](Word.WdHelpType.md)** constants.|

## Remarks

Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.


## Example

This example displays the Help Topics dialog box.


```vb
Help HelpType:=wdHelp
```

This example displays a list of Help topics that describe how to use Help.




```vb
Help HelpType:=wdHelpUsingHelp
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]