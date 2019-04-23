---
title: TextRange2.Replace method (Office)
ms.prod: office
api_name:
- Office.TextRange2.Replace
ms.assetid: e14f0ad0-3b9c-d9f5-a13d-d3bbdcae50e1
ms.date: 01/25/2019
localization_priority: Normal
---


# TextRange2.Replace method (Office)

Finds specific text in a text range, replaces the found text with a specified string, and returns a **TextRange2** object that represents the first occurrence of the found text. Returns **Nothing** if no match is found.


## Syntax

_expression_.**Replace** (_FindWhat_, _ReplaceWhat_, _After_, _MatchCase_, _WholeWords_)

_expression_ An expression that returns a **[TextRange2](Office.TextRange2.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FindWhat_|Required|**String**|The text to search for.|
| _ReplaceWhat_|Required|**String**|The text you want to replace the found text with.|
| _After_|Optional|**Long**|The position of the character (in the specified text range) after which you want to search for the next occurrence of **FindWhat**. For example, if you want to search from the fifth character of the text range, specify 4 for **After**. If this argument is omitted, the first character of the text range is used as the starting point for the search.|
| _MatchCase_|Optional|**[MsoTriState](office.msotristate.md)**|Determines whether a distinction is made on the basis of case.|
| _WholeWords_|Optional|**MsoTriState**|Determines whether only whole words are searched.|

## Return value

TextRange2


## See also

- [TextRange2 object members](overview/Library-Reference/textrange2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]