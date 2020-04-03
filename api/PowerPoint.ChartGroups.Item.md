---
title: ChartGroups.Item method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroups.Item
ms.assetid: 0b04a471-d726-f400-062c-8d4a7dc9c752
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroups.Item method (PowerPoint)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a '[ChartGroups](PowerPoint.ChartGroups.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The index number for the object.|

## Return value

A  **[ChartGroup](PowerPoint.ChartGroup.md)** object contained by the collection.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds drop lines to chart group one for the first chart group of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartGroups.Item(1).HasDropLines = True

    End If

End With
```


## See also


[ChartGroups Object](PowerPoint.ChartGroups.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]