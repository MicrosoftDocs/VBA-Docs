---
title: TextRange2.PasteSpecial method (PowerPoint)
ms.assetid: 05855fac-1123-44dd-b021-553216485db6
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# TextRange2.PasteSpecial method (PowerPoint)

Replaces the text range with the contents of the Clipboard in the format specified. If the paste succeeds, this method returns a **TextRange2** object including the text range that was pasted.


## Syntax

_expression_.**PasteSpecial** (_Format_)

 _expression_ An expression that returns a 'TextRange2' object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Format_|Required|**MsoClipboardFormat**|Determines the format for the Clipboard contents when they're inserted into the document.|

## Return value

TextRange2


## See also


[TextRange2 object (PowerPoint)](PowerPoint.textrange2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]