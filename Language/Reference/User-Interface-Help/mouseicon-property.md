---
title: MouseIcon property 
keywords: fm20.chm2001540
f1_keywords:
- fm20.chm2001540
ms.prod: office
api_name:
- Office.MouseIcon
ms.assetid: b5834d6d-76ad-73e6-b55d-0ab4caa643ef
ms.date: 11/16/2018
localization_priority: Normal
---


# MouseIcon property 

Assigns a custom icon to an object.

## Syntax

_object_.**MouseIcon** = **LoadPicture(**_pathname_**)**

The **MouseIcon** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _pathname_|Required. A string expression specifying the path and filename of the file containing the custom icon.|

## Remarks

The **MouseIcon** property is valid when the **MousePointer** property is set to 99. The mouse icon of an object is the image that appears when the user moves the mouse across that object.

To assign an image for the mouse pointer, you can either assign a picture to the **MouseIcon** property or load a picture from a file by using the **LoadPicture** function.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]