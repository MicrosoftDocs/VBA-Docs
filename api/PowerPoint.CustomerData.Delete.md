---
title: CustomerData.Delete method (PowerPoint)
keywords: vbapp10.chm675005
f1_keywords:
- vbapp10.chm675005
ms.prod: powerpoint
api_name:
- PowerPoint.CustomerData.Delete
ms.assetid: 7a7649f9-7efa-57e7-15db-a16991dc6f09
ms.date: 06/08/2017
localization_priority: Normal
---


# CustomerData.Delete method (PowerPoint)

Deletes the specified  **[CustomXMLPart](Office.CustomXMLPart.md)** object from the **[CustomerData](PowerPoint.CustomerData.md)** collection of a **[CustomLayout](PowerPoint.CustomLayout.md)**, **[Master](PowerPoint.Master.md)**, **[Presentation](PowerPoint.Presentation.md)**, **[Shape](PowerPoint.Shape.md)**, or **[Slide](PowerPoint.Slide.md)** object.


## Syntax

_expression_.**Delete**( `_Id_` )

 _expression_ An expression that returns a [CustomerData](PowerPoint.CustomerData.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Id_|Required|**String**|The ID of the  **CustomXMLPart** object to be deleted.|

## Remarks

Individual  **CustomXMLPart** objects in the **CustomerData** collection are represented by GUIDs (globally unique identifiers). Pass the GUID that represents the custom XML part that you want to delete to the Id parameter of the **Delete** method as a **String**. You can get the ID of a particular custom XML part by iterating through the collection, using the **Id** property of the **CustomerData** collection.


## Example

The following example shows how to use the Delete method to delete a custom XML part from the  **CustomerData** collection. It adds a new custom XML part to the **CustomerData** collection of the first shape on the first slide of the active presentation. Then it gets the ID of the new part and passes it to the **Delete** method.


```vb
Public Sub Delete_Example() 
 
    Dim pptSlide As Slide 
    Set pptSlide = ActivePresentation.Slides(1) 
     
    Dim pptShape As Shape 
    Set pptShape = pptSlide.Shapes(1) 
     
    ' Get the CustomerData collection of the shape 
    Dim pptCustomerData As customerData 
    Set pptCustomerData = pptShape.customerData 
     
    ' Get the current count of custom XML parts 
    Debug.Print pptCustomerData.Count 
     
    ' Add a new CustomXMLPart object to the CustomerData 
    ' collection for this shape and get the revised count of 
    ' custom XML parts 
    Dim pptCustomXMLPart As CustomXMLPart 
    Set pptCustomXMLPart = pptCustomerData.Add 
    Debug.Print pptCustomerData.Count 
     
    ' Get the ID of the new part 
    Dim myString As String 
    myString = pptCustomXMLPart.Id 
    Debug.Print myString 
     
    ' Delete the new part and re-check the count of custom XML parts 
    pptCustomerData.Delete (myString) 
    Debug.Print pptCustomerData.Count 
 
End Sub
```


## See also


[CustomerData Collection](PowerPoint.CustomerData.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]