---
title: CalloutFormat.DropType property (Publisher)
keywords: vbapb10.chm2490630
f1_keywords:
- vbapb10.chm2490630
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.DropType
ms.assetid: fd4ec192-0732-e860-4ff8-e305aa0d90a9
ms.date: 06/05/2019
localization_priority: Normal
---


# CalloutFormat.DropType property (Publisher)

Returns an **[MsoCalloutDropType](Office.MsoCalloutDropType.md)** constant indicating where the callout line attaches to the callout text box. Read-only.


## Syntax

_expression_.**DropType**

_expression_ A variable that represents a **[CalloutFormat](Publisher.CalloutFormat.md)** object.


## Return value

MsoCalloutDropType


## Remarks

The **DropType** property value can be one of the **MsoCalloutDropType** constants declared in the Microsoft Office type library.

If the callout drop type is **msoCalloutDropCustom**, the values of the **[Drop](Publisher.CalloutFormat.Drop.md)** and **[AutoAttach](Publisher.CalloutFormat.AutoAttach.md)** properties and the relative positions of the callout text box and callout line origin (where the callout points) are used to determine where the callout line attaches to the text box.

Use the **[PresetDrop](Publisher.CalloutFormat.PresetDrop.md)** method to set the value of this property.


## Example

This example replaces the custom drop for the first shape in the active publication with one of two preset drops, depending on whether the custom drop value is greater than or less than half the height of the callout text box. For the example to work, the shape must be a callout.

```vb
With ActiveDocument.Pages(1).Shapes(1).Callout 
 If .DropType = msoCalloutDropCustom Then 
 If .Drop < .Parent.Height / 2 Then 
 .PresetDrop DropType:=msoCalloutDropTop 
 Else 
 .PresetDrop DropType:=msoCalloutDropBottom 
 End If 
 End If 
End With 

```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]