---
title: TextRange2.InsertSymbol method (Office)
ms.prod: office
api_name:
- Office.TextRange2.InsertSymbol
ms.assetid: 74642859-0d84-5de9-494a-a58b9d93de88
ms.date: 01/25/2019
localization_priority: Normal
---


# TextRange2.InsertSymbol method (Office)

Inserts a symbol from the specified font set into the range of text represented by the **TextRange2** object.


## Syntax

_expression_.**InsertSymbol** (_FontName_, _CharNumber_, _Unicode_)

_expression_ An expression that returns a **[TextRange2](Office.TextRange2.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FontName_|Required|**String**|The name of the font set.|
| _CharNumber_|Required|**Long**|The number of the symbol.|
| _Unicode_|Optional|**[MsoTriState](office.msotristate.md)**|Indicates whether the value of the symbol is specified as a unicode value.|

## Return value

TextRange2


## See also

- [TextRange2 object members](overview/Library-Reference/textrange2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]