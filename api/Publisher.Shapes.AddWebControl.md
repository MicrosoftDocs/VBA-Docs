---
title: Shapes.AddWebControl method (Publisher)
keywords: vbapb10.chm2162722
f1_keywords:
- vbapb10.chm2162722
ms.prod: publisher
api_name:
- Publisher.Shapes.AddWebControl
ms.assetid: 94b54939-9627-6b38-4375-f1c87fc8c4f7
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddWebControl method (Publisher)

Adds a new **[Shape](Publisher.Shape.md)** object representing a web form control to the specified **Shapes** collection.


## Syntax

_expression_.**AddWebControl** (_Type_, _Left_, _Top_, _Width_, _Height_, _LaunchPropertiesWindow_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_ |Required| **[PbWebControlType](publisher.pbwebcontroltype.md)**|Specifies the type of web form control to add. Can be one of the **PbWebControlType** constants declared in the Microsoft Publisher type library. An error occurs if **pbWebControlWebComponent** is used.|
| _Left_ |Required| **Variant**|The position of the left edge of the shape representing the web form control.|
| _Top_ |Required| **Variant**|The position of the top edge of the shape representing the web form control.|
| _Width_|Required| **Variant**|The width of the shape representing the web form control. For command buttons, this parameter is ignored.|
| _Height_|Required| **Variant**|The height of the shape representing the web form control. For command buttons, this parameter is ignored.|
|_LaunchPropertiesWindow_|Optional| **Boolean**|Not supported. Default is **False**; an error occurs if this argument is set to **True**.|

## Return value

Shape


## Remarks

For the _Left_, _Top_, _Width_, and _Height_ parameters, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

When adding a hot spot to a web control by using the **pbWebControlHotSpot** constant, the URL is specified by the **[Hyperlinks](Publisher.TextRange.Hyperlinks.md)** property.

Note that the **Shape.Fill** property, which returns a **FillFormat** object, and the **Shape.Line** property, which returns a **LineFormat** object, cannot be accessed from a hot spot shape. A run-time error is returned if attempting to access these properties from a hot spot shape.

## Example

The following example adds a web form check box control to the first page of the active publication.

```vb
Dim shpCheckBox As Shape 
 
Set shpCheckBox = ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCheckBox, _ 
 Left:=216, Top:=216, _ 
 Width:=18, Height:=18) 

```

<br/>

The following example adds hot spots to a shape on page four of the active web publication. First, a four-point star AutoShape is added to the page. Next, a hot spot is added to each arm of the star by using the **AddWebControl** method with a _Type_ of **pbWebControlHotSpot**. Finally, a hyperlink is added to each hot spot by using the **Hyperlinks** property of each hot spot shape.

```vb
Dim theDoc As Document 
Dim theStar As Shape 
Dim theWC1 As Shape 
Dim theWC2 As Shape 
Dim theWC3 As Shape 
Dim theWC4 As Shape 
 
Set theDoc = ActiveDocument 
Set theStar = theDoc.Pages(4).Shapes.AddShape _ 
 (Type:=msoShape4pointStar, Left:=200, Top:=25, _ 
 Width:=200, Height:=200) 
 
With theDoc.Pages(4).Shapes 
 
 Set theWC1 = .AddWebControl(Type:=pbWebControlHotSpot, _ 
 Left:=280, Top:=25, Width:=40, Height:=80) 
 With theWC1 
 .Hyperlink.Address = "https://www.contoso.com/page1.htm" 
 End With 
 
 Set theWC2 = .AddWebControl(Type:=pbWebControlHotSpot, _ 
 Left:=320, Top:=105, Width:=80, Height:=40) 
 With theWC2 
 .Hyperlink.Address = "https://www.contoso.com/page2.htm" 
 End With 
 
 Set theWC3 = .AddWebControl(Type:=pbWebControlHotSpot, _ 
 Left:=280, Top:=145, Width:=40, Height:=80) 
 With theWC3 
 .Hyperlink.Address = "https://www.contoso.com/page3.htm" 
 End With 
 
 Set theWC4 = .AddWebControl(Type:=pbWebControlHotSpot, _ 
 Left:=200, Top:=105, Width:=80, Height:=40) 
 With theWC4 
 .Hyperlink.Address = "https://www.contoso.com/page4.htm" 
 End With 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]