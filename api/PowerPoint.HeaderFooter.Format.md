---
title: HeaderFooter.Format property (PowerPoint)
keywords: vbapp10.chm582006
f1_keywords:
- vbapp10.chm582006
ms.prod: powerpoint
api_name:
- PowerPoint.HeaderFooter.Format
ms.assetid: ba8f2afa-8c57-60e0-cd84-9366c016efd9
ms.date: 06/08/2017
localization_priority: Normal
---


# HeaderFooter.Format property (PowerPoint)

Returns or sets the format for the automatically updated date and time. Read/write.


## Syntax

_expression_.**Format**

_expression_ A variable that represents a [ThreeDFormat](PowerPoint.ThreeDFormat.md) object.


## Return value

PpDateTimeFormat


## Remarks

The **Format** property applies only to **HeaderFooter** objects that represent a date and time (returned from the **HeadersFooters** collection by the **[DateAndTime](PowerPoint.HeadersFooters.DateAndTime.md)** property).

Make sure that the date and time are set to be updated automatically (not displayed as fixed text) by setting the  **[UseFormat](PowerPoint.HeaderFooter.UseFormat.md)** property to **True**.

The **Format** property value can be one of these **PpDateTimeFormat** constants.


||
|:-----|
|**ppDateTimeddddMMMMddyyyy**|
|**ppDateTimedMMMMyyyy**|
|**ppDateTimedMMMyy**|
|**ppDateTimeFormatMixed**|
|**ppDateTimeHmm**|
|**ppDateTimehmmAMPM**|
|**ppDateTimeHmmss**|
|**ppDateTimehmmssAMPM**|
|**ppDateTimeMdyy**|
|**ppDateTimeMMddyyHmm**|
|**ppDateTimeMMddyyhmmAMPM**|
|**ppDateTimeMMMMdyyyy**|
|**ppDateTimeMMMMyy**|
|**ppDateTimeMMyy**|

## Example

This example sets the date and time for the slide master of the active presentation to be updated automatically and then it sets the date and time format to show hours, minutes, and seconds.


```vb
Set myPres = Application.ActivePresentation

With myPres.SlideMaster.HeadersFooters.DateAndTime

    .UseFormat = True

    .Format = ppDateTimeHmmss

End With
```


## See also


[HeaderFooter Object](PowerPoint.HeaderFooter.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]