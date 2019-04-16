---
title: CalloutFormat.Accent property (Excel)
keywords: vbaxl10.chm104006
f1_keywords:
- vbaxl10.chm104006
ms.prod: excel
api_name:
- Excel.CalloutFormat.Accent
ms.assetid: 9dce6821-47df-174d-c7f3-7edad9fcf77d
ms.date: 04/13/2019
localization_priority: Normal
---


# CalloutFormat.Accent property (Excel)

Allows the user to place a vertical accent bar to separate the callout text from the callout line. Read/write **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**Accent**

_expression_ A variable that represents a **[CalloutFormat](Excel.CalloutFormat.md)** object.


## Remarks

A vertical accent bar separates the callout text from the callout line.

## Example

This example adds to _myDocument_ an oval and a callout that points to the oval. The callout text won't have a border, but it will have a vertical accent bar that separates the text from the callout line.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
    .AddShape msoShapeOval, 180, 200, 280, 130 
    With .AddCallout(msoCalloutTwo, 420, 170, 170, 40) 
        .TextFrame.Characters.Text = "My oval" 
        With .Callout 
            .Accent = msoTrue 
            .Border = False 
        End With 
    End With 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]