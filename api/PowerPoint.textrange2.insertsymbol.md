---
title: TextRange2.InsertSymbol method (PowerPoint)
ms.assetid: cfec2f5d-fcd8-4a49-bf1e-5c86a0570ff7
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# TextRange2.InsertSymbol method (PowerPoint)

Inserts a symbol from the specified font set into the range of text represented by the  **TextRange2** object.


## Syntax

_expression_. `InsertSymbol`( `_FontName_`, `_CharNumber_`, `_Unicode_` )

 _expression_ An expression that returns a 'TextRange2' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FontName_|Required|**String**|The name of the font set.|
| _CharNumber_|Required|**Long**|The number of the symbol.|
| _Unicode_|Optional|**MsoTriState**|Indicates whether the value of the symbol is specified as a unicode value.|

## Return value

TextRange2


## See also


[TextRange2 object (PowerPoint)](PowerPoint.textrange2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]