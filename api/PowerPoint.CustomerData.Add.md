---
title: CustomerData.Add method (PowerPoint)
keywords: vbapp10.chm675004
f1_keywords:
- vbapp10.chm675004
ms.prod: powerpoint
api_name:
- PowerPoint.CustomerData.Add
ms.assetid: f39bc83a-4c3b-6803-12d1-9ae72e601b49
ms.date: 06/08/2017
localization_priority: Normal
---


# CustomerData.Add method (PowerPoint)

 Adds a **[CustomXMLPart](Office.CustomXMLPart.md)** to the **[CustomerData](PowerPoint.CustomerData.md)** collection of a **[CustomLayout](PowerPoint.CustomLayout.md)**, **[Master](PowerPoint.Master.md)**, **[Presentation](PowerPoint.Presentation.md)**, **[Shape](PowerPoint.Shape.md)**, or **[Slide](PowerPoint.Slide.md)** object and returns the **CustomXMLPart** object created.


## Syntax

_expression_.**Add**

 _expression_ An expression that returns a [CustomerData](PowerPoint.CustomerData.md) object.


## Return value

CustomXMLPart


## Remarks

You can add one or more items of customer data (custom XML parts) to any of the objects listed above that can contain customer data.


## Example




```vb
Public Sub Add_Example() 
 
    Dim pptSlide As Slide 
    Set pptSlide = ActivePresentation.Slides(1) 
     
    Dim pptShape As Shape 
    For Each pptShape In pptSlide.Shapes 
         
        ' Get the CustomerData collection of the shape 
        Dim pptCustomerData As customerData 
        Set pptCustomerData = pptShape.customerData 
         
        ' Add a new CustomXMLPart object to the CustomerData collection for this shape 
        Dim pptCustomXMLPart As CustomXMLPart 
        Set pptCustomXMLPart = pptCustomerData.Add 
         
        ' Add data to the CustomXMLPart 
        pptCustomXMLPart.LoadXML ("<ShapeData><DataItem>This has to be valid XML</DataItem></ShapeData>") 
         
        ' Print the ID (a GUID) of the CustomXMLPart 
        Debug.Print (pptCustomXMLPart.Id) 
         
    Next 
 
End Sub
```


## See also


[CustomerData Collection](PowerPoint.CustomerData.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]