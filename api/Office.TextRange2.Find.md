---
title: TextRange2.Find Method (Office)
ms.prod: office
api_name:
- Office.TextRange2.Find
ms.assetid: ad5bc61a-a7f1-485a-0fc8-a3bd6707f956
ms.date: 06/08/2017
---


# TextRange2.Find Method (Office)

Searches a  **TextRange2** object for a subset of text.


## Syntax

 _expression_. `Find`( `_FindWhat_`, `_After_`, `_MatchCase_`, `_WholeWords_` )

 _expression_ An expression that returns a [TextRange2](./Office.TextRange2.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FindWhat_|Required|**String**|Contains the text to find.|
| _After_|Optional|**Long**|Specifies the point in the text range to start the search.|
| _MatchCase_|Optional|**MsoTriState**|Specifies if the target text must exactly match the case of the search text. |
| _WholeWords_|Optional|**MsoTriState**|Specifies that only whole words will be searched.|

### Return value

TextRange2


## See also


[TextRange2 Object](Office.TextRange2.md)



[TextRange2 Object Members](./overview/Library-Reference/textrange2-members-office.md)

