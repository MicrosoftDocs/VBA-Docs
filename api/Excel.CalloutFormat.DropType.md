---
title: CalloutFormat.DropType property (Excel)
keywords: vbaxl10.chm104012
f1_keywords:
- vbaxl10.chm104012
ms.prod: excel
api_name:
- Excel.CalloutFormat.DropType
ms.assetid: ab947fa4-4af9-e491-f62d-e0ca036e1892
ms.date: 04/13/2019
localization_priority: Normal
---


# CalloutFormat.DropType property (Excel)

Returns a value that indicates where the callout line attaches to the callout text box. Read-only **[MsoCalloutDropType](Office.MsoCalloutDropType.md)**.


## Syntax

_expression_.**DropType**

_expression_ A variable that represents a **[CalloutFormat](Excel.CalloutFormat.md)** object.


## Remarks

If the callout drop type is **msoCalloutDropCustom**, the values of the **[Drop](Excel.CalloutFormat.Drop.md)** and **[AutoAttach](Excel.CalloutFormat.AutoAttach.md)** properties and the relative positions of the callout text box and callout line origin (the place that the callout points to) are used to determine where the callout line attaches to the text box.

This property is read-only. Use the **[PresetDrop](Excel.CalloutFormat.PresetDrop.md)** method to set the value of this property.


## Example

This example replaces the custom drop for shape one on _myDocument_ with one of two preset drops, depending on whether the custom drop value is greater than or less than half the height of the callout text box. For the example to work, shape one must be a callout.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Callout 
    If .DropType = msoCalloutDropCustom Then 
        If .Drop < .Parent.Height / 2 Then 
            .PresetDrop msoCalloutDropTop 
        Else 
            .PresetDrop msoCalloutDropBottom 
        End If 
    End If 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]