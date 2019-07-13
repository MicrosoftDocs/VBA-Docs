---
title: TextStyles.Item method (PowerPoint)
keywords: vbapp10.chm578003
f1_keywords:
- vbapp10.chm578003
ms.prod: powerpoint
api_name:
- PowerPoint.TextStyles.Item
ms.assetid: 3315d566-a46a-38cc-44b3-07c54ec3c6e5
ms.date: 06/08/2017
localization_priority: Normal
---


# TextStyles.Item method (PowerPoint)

Returns a single text style from the specified  **[TextStyles](PowerPoint.TextStyles.md)** collection.


## Syntax

_expression_.**Item** (_Type_)

_expression_ A variable that represents a [TextStyles](PowerPoint.TextStyles.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**PpTextStyleType**|The text style type.|

## Return value

TextStyle


## Remarks

The  **Item** method is the default member for a collection. For example, the following two lines of code are equivalent:

 `ActivePresentation.Slides.Item(1)`

 `ActivePresentation.Slides(1)`

The  _Type_ parameter value can be one of these **PpTextStyleType** constants.


||
|:-----|
|**ppBodyStyle**|
|**ppDefaultStyle**|
|**ppTitleStyle**|

## See also


[TextStyles Object](PowerPoint.TextStyles.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]