---
title: CalloutFormat.AutoLength property (PowerPoint)
keywords: vbapp10.chm559009
f1_keywords:
- vbapp10.chm559009
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.AutoLength
ms.assetid: 40578d3b-b23d-cf11-51a0-d59c3cf2a226
ms.date: 06/08/2017
localization_priority: Normal
---


# CalloutFormat.AutoLength property (PowerPoint)

Determines whether the first segment of the callout retains the fixed length specified by the  **[Length](PowerPoint.CalloutFormat.Length.md)** property, or is scaled automatically, whenever the callout is moved. Read-only.


## Syntax

_expression_.**AutoLength**

_expression_ A variable that represents an [CalloutFormat](PowerPoint.CalloutFormat.md) object.


## Return value

MsoTriState


## Remarks

This property is read-only. However, you can use the  **[AutomaticLength](PowerPoint.CalloutFormat.AutomaticLength.md)** method to set this property to **msoTrue** and the **[CustomLength](PowerPoint.CalloutFormat.CustomLength.md)** method to set this property to **msoFalse**.

The value returned by the  **AutoLength** property can be one of these **MsoTriState** constants. This property applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour**).



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The first segment of the callout retains the fixed length specified by the  **Length** property whenever the callout is moved.|
|**msoTrue**| The first segment of the callout line (the segment attached to the text callout box) is scaled automatically whenever the callout is moved.|

## Example

This example switches between an automatically scaling first segment and one with a fixed length for the callout line for shape one on _myDocument_. For the example to work, shape one must be a callout.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).Callout

    If .AutoLength Then

        .CustomLength 50

    Else

        .AutomaticLength

    End If

End With
```


## See also


[CalloutFormat Object](PowerPoint.CalloutFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]