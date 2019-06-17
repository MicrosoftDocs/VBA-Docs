---
title: Shapes.AddWebNavigationBar method (Publisher)
keywords: vbapb10.chm2162736
f1_keywords:
- vbapb10.chm2162736
ms.prod: publisher
api_name:
- Publisher.Shapes.AddWebNavigationBar
ms.assetid: 26e9622c-ea28-b28b-9904-b3a3ccc9341b
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddWebNavigationBar method (Publisher)

Adds a **[Shape](Publisher.Shape.md)** object of type **pbWebNavigationBar** (**[PbShapeType](publisher.pbshapetype.md)** enumeration) to the current page of a publication.


## Syntax

_expression_.**AddWebNavigationBar** (_Name_, _Left_, _Top_, _Width_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Name_|Required| **String**|The name of the **[WebNavigationBarSet](publisher.webnavigationbarset.md)** object to add to the specified **Shape**.|
|_Left_ |Required| **Variant**|The position of the left edge of the shape that represents the web navigation bar set.|
|_Top_ |Required| **Variant**|The position of the top edge of the shape that represents the web navigation bar set.|
|_Width_|Optional| **Variant**|The width of the shape that represents the web navigation bar set.|

## Return value

Shape


## Remarks

The **AddWebNavigationBar** method does not create a web navigation bar set. It adds an existing set from the **WebNavigationBarSets** collection. Pass the name of the existing web navigation bar set as the _Name_ parameter.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **AddWebNavigationBar** method to add a **WebNavigationBarSet** object to the active document.

```vb
Public Sub AddWebNavigationBarSet_Example() 
 
 Dim pubShape As Publisher.Shape 
 
 ThisDocument.WebNavigationBarSets.AddSet ("NavBar") 
 Set pubShape = ThisDocument.Pages(1).Shapes.AddWebNavigationBar("NavBar", 10, 25) 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]