---
title: HeaderFooter.UseFormat property (PowerPoint)
keywords: vbapp10.chm582005
f1_keywords:
- vbapp10.chm582005
ms.prod: powerpoint
api_name:
- PowerPoint.HeaderFooter.UseFormat
ms.assetid: da9739ea-fb9b-5e3d-bb7e-64763ef11bf2
ms.date: 06/08/2017
localization_priority: Normal
---


# HeaderFooter.UseFormat property (PowerPoint)

Determines whether the date and time object contains automatically updated information. Read/write.


## Syntax

_expression_. `UseFormat`

_expression_ A variable that represents an [HeaderFooter](PowerPoint.HeaderFooter.md) object.


## Return value

MsoTriState


## Remarks

This property applies only to a  **[HeaderFooter](PowerPoint.HeaderFooter.md)** object that represents a date and time (returned by the **[DateAndTime](PowerPoint.HeadersFooters.DateAndTime.md)** property). Set the **UseFormat** property of a date and time **HeaderFooter** object to **True** when you want to set or return the date and time format by using the **[Format](PowerPoint.HeaderFooter.Format.md)** property. Set the **UseFormat** property to **msoFalse** when you want to set or return the text string for the fixed date and time.

The value of the  **UseFormat** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The date and time object is a fixed string.|
|**msoTrue**| The date and time object contains automatically updated information.|

## Example

This example sets the date and time for the slide master of the active presentation to be updated automatically and then it sets the date and time format to show hours, minutes, and seconds.


```vb
Set myPres = Application.ActivePresentation

With myPres.SlideMaster.HeadersFooters.DateAndTime

    .UseFormat = msoTrue

    .Format = ppDateTimeHmmss

End With
```


## See also


[HeaderFooter Object](PowerPoint.HeaderFooter.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]