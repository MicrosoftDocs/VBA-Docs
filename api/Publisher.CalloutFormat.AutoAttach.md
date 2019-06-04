---
title: CalloutFormat.AutoAttach property (Publisher)
keywords: vbapb10.chm2490626
f1_keywords:
- vbapb10.chm2490626
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.AutoAttach
ms.assetid: 893303d8-97fe-9eea-8d6e-d9110c75ee84
ms.date: 06/05/2019
localization_priority: Normal
---


# CalloutFormat.AutoAttach property (Publisher)

Returns or sets an **[MsoTriState](Office.MsoTriState.md)** constant indicating whether the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line (where the callout points) is to the left or right of the callout text box. Read/write.


## Syntax

_expression_.**AutoAttach**

_expression_ A variable that represents a **[CalloutFormat](Publisher.CalloutFormat.md)** object.


## Return value

MsoTriState


## Remarks

The **AutoAttach** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library.

When the value of this property is **msoTrue**, the drop value (the vertical distance from the edge of the callout text box to the place where the callout line attaches) is measured from the top of the text box when the text box is to the right of the origin, and it is measured from the bottom of the text box when the text box is to the left of the origin. 

When the value of this property is **msoFalse**, the drop value is always measured from the top of the text box, regardless of the relative positions of the text box and the origin. Use the **[CustomDrop](Publisher.CalloutFormat.CustomDrop.md)** method to set the drop value, and use the **[Drop](Publisher.CalloutFormat.Drop.md)** property to return the drop value.

Setting this property affects a callout only if it has an explicitly set drop value—that is, if the value of the **[DropType](Publisher.CalloutFormat.DropType.md)** property is **msoCalloutDropCustom**. By default, callouts have explicitly set drop values when they are created.


## Example

This example adds two callouts to the first page. One of the callouts is automatically attached and the other is not. If you change the callout line origin for the automatically attached callout to the right of the attached text box, the position of the text box changes. The callout that is not automatically attached does not display this behavior.

```vb
With ActivePublication.Pages(1).Shapes 
 With .AddCallout(Type:=msoCalloutTwo, _ 
 Left:=420, Top:=170, Width:=200, Height:=50) 
 .TextFrame.TextRange.Text = "auto-attached" 
 .Callout.AutoAttach = msoTrue 
 End With 
 With .AddCallout(Type:=msoCalloutTwo, _ 
 Left:=420, Top:=350, Width:=200, Height:=50) 
 .TextFrame.TextRange.Text = "not auto-attached" 
 .Callout.AutoAttach = msoFalse 
 End With 
End With 

```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]