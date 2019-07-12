---
title: TextEffectFormat.FontItalic property (PowerPoint)
keywords: vbapp10.chm556005
f1_keywords:
- vbapp10.chm556005
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.FontItalic
ms.assetid: ee7b38b5-2ef7-ba05-e986-b3c84881baed
ms.date: 06/08/2017
localization_priority: Normal
---


# TextEffectFormat.FontItalic property (PowerPoint)

Determines whether the font in the specified WordArt is italic. Read/write.


## Syntax

_expression_. `FontItalic`

_expression_ A variable that represents a [TextEffectFormat](PowerPoint.TextEffectFormat.md) object.


## Return value

MsoTriState


## Remarks

The value of the  **FontItalic** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The font in the specified WordArt is not italic.|
|**msoTrue**| The font in the specified WordArt is italic.|

## Example

This example sets the font to italic for the shape named "WordArt 4" on _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes("WordArt 4").TextEffect.FontItalic = msoTrue
```


## See also


[TextEffectFormat Object](PowerPoint.TextEffectFormat.md)
[Font Object](PowerPoint.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]