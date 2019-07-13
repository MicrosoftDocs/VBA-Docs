---
title: Font.Bold property (PowerPoint)
keywords: vbapp10.chm575004
f1_keywords:
- vbapp10.chm575004
ms.prod: powerpoint
api_name:
- PowerPoint.Font.Bold
ms.assetid: 13e81c46-5ae7-21ee-58e1-5ab23de552d5
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.Bold property (PowerPoint)

Determines whether the character format is bold. Read/write.


## Syntax

_expression_.**Bold**

_expression_ A variable that represents a [Font](PowerPoint.Font.md) object.


## Return value

MsoTriState


## Remarks

The value of the  **Bold** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The character format is not bold.|
|**msoTriStateMixed**|The specified text range contains both bold and nonbold characters.|
|**msoTrue**| The character format is bold.|

## Example

This example sets characters one through five in the title on slide one to bold.


```vb
Set myT = Application.ActivePresentation.Slides(1).Shapes.Title

myT.TextFrame.TextRange.Characters(1, 5).Font.Bold = msoTrue
```


## See also


[Font Object](PowerPoint.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]