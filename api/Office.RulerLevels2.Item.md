---
title: RulerLevels2.Item method (Office)
ms.prod: office
api_name:
- Office.RulerLevels2.Item
ms.assetid: b6791181-ea32-62e3-3b9a-1b60f436bc91
ms.date: 01/23/2019
localization_priority: Normal
---


# RulerLevels2.Item method (Office)

Gets a member of the **RulerLevels2** collection.


## Syntax

_expression_.**Item**(_Index_)

_expression_ An expression that returns a **[RulerLevels2](Office.RulerLevels2.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The index number of the object to be returned.|

## Return value

RulerLevel2


## Example

This example sets the first-line indent and the hanging indent for outline level one in body text on the slide master for the active presentation.


```vb
With ActivePresentation.SlideMaster.TextStyles.Item(ppBodyStyle) 
 With .Ruler2.Levels.Item(1) ' sets indents for level 1 
 .FirstMargin = 9 
 .LeftMargin = 54 
 End With 
End With 

```


## See also

- [RulerLevels2 object members](overview/Library-Reference/rulerlevels2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]