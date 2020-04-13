---
title: Selection.InsertCaption method (Word)
keywords: vbawd10.chm158663073
f1_keywords:
- vbawd10.chm158663073
ms.prod: word
api_name:
- Word.Selection.InsertCaption
ms.assetid: 848c1686-ca8c-d022-68f1-74a2f3d46498
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.InsertCaption method (Word)

Inserts a caption immediately preceding or following the specified selection.


## Syntax

_expression_. `InsertCaption`( `_Label_` , `_Title_` , `_TitleAutoText_` , `_Position_` , `_ExcludeLabel_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Label_|Required| **Variant**|The caption label to be inserted. Can be a  **String** or one of the **WdCaptionLabelID** constants. If the label has not yet been defined, an error occurs. Use the **Add** method with the **CaptionLabels** object to define new caption labels.|
| _Title_|Optional| **Variant**|The string to be inserted immediately following the label in the caption (ignored if TitleAutoText is specified).|
| _TitleAutoText_|Optional| **Variant**|The AutoText entry whose contents you want to insert immediately following the label in the caption (overrides any text specified by Title).|
| _Position_|Optional| **Variant**|Specifies whether the caption will be inserted above or below the selection. Can be one of the **WdCaptionPosition** constants.|
| _ExcludeLabel_|Optional| **Variant**| **True** does not include the text label, as defined in the Label parameter. **False** includes the specified label.|

## Example

This example inserts a Figure caption at the insertion point.


```vb
Selection.Collapse Direction:=wdCollapseStart 
Selection.InsertCaption Label:="Figure", _ 
 Title:=": Sales Results", Position:=wdCaptionPositionBelow
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]