---
title: Shape.Data1 Property (Visio)
keywords: vis_sdr.chm11213365
f1_keywords:
- vis_sdr.chm11213365
ms.prod: visio
api_name:
- Visio.Shape.Data1
ms.assetid: ca9dda75-4ae2-70f0-46bd-ff5afbba84fc
ms.date: 06/08/2017
---


# Shape.Data1 Property (Visio)

Gets or sets the value of the  **Data1** field for a **Shape** object. Read/write.


## Syntax

 _expression_. `Data1`

 _expression_ A variable that represents a [Shape](./Visio.Shape.md) object.


### Return value

String


## Remarks

Use the  **Data1** property to supply additional information about a shape. The property can contain up to 64 KB of characters. Text controls should be used with care with a string that is greater than 3,000 characters. Setting the **Data1** property is equivalent to entering information in the **Data 1** box in the **Special** dialog box (click **Shape Name** in the **Shape Design** group on the [Developer](../visio/How-to/run-visio-in-developer-mode.md) tab).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to set a shape's  **Data1** , **Data2** , and **Data3** properties. It prints the values of these properties in the **Immediate** window. You can also verify that these values have been set by opening the **Special** dialog box.


```vb
Public Sub Data123_Example() 
 
 Dim vsoPage As Visio.Page 
 Dim vsoShape As Visio.Shape 
 
 Set vsoPage = Documents.Add("").Pages(1) 
 Set vsoShape = vsoPage.DrawRectangle(3, 3, 5, 5) 
 
 'Use the Data1, Data2, and Data3 properties to set 
 'the shape's Data fields. 
 vsoShape.Data1 = "Data1_String" 
 vsoShape.Data2 = "Data2_String" 
 vsoShape.Data3 = "Data3_String" 
 
 'Use the Data1, Data2, and Data3 properties to verify 
 'the shape's Data field values. 
 Debug.Print vsoShape.Data1 
 Debug.Print vsoShape.Data2 
 Debug.Print vsoShape.Data3 
 
End Sub
```


