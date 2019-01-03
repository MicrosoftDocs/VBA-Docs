---
title: TextRange2.PasteSpecial method (Office)
ms.prod: office
api_name:
- Office.TextRange2.PasteSpecial
ms.assetid: 79f88454-2f95-ea10-6ec4-5fb78ca8036d
ms.date: 06/08/2017
---


# TextRange2.PasteSpecial method (Office)

Replaces the text range with the contents of the Clipboard in the format specified. If the paste succeeds, this method returns a  **TextRange2** object including the text range that was pasted.


## Syntax

 _expression_. `PasteSpecial`( `_Format_` )

 _expression_ An expression that returns a [TextRange2](Office.TextRange2.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Format_|Required|**MsoClipboardFormat**|Determines the format for the Clipboard contents when they're inserted into the document.|

## Return value

TextRange2


## See also


[TextRange2 Object](Office.TextRange2.md)



[TextRange2 Object Members](./overview/Library-Reference/textrange2-members-office.md)

