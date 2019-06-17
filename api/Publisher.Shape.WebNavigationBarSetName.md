---
title: Shape.WebNavigationBarSetName property (Publisher)
keywords: vbapb10.chm5308677
f1_keywords:
- vbapb10.chm5308677
ms.prod: publisher
api_name:
- Publisher.Shape.WebNavigationBarSetName
ms.assetid: 0d9abe17-6936-562b-9210-5f092d13f215
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.WebNavigationBarSetName property (Publisher)

Returns a **String** that represents the name of the web navigation bar set that the specified shape is an instance of. Read-only.


## Syntax

_expression_.**WebNavigationBarSetName**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Return value

String


## Remarks

This property is only accessible for shapes that represent an instance of a web navigation bar set. Use the **[Type](Publisher.Shape.Type.md)** property to determine if a shape represents an instance of a web navigation bar set.

Use the **WebNavigationBarSetName** property to return the name of a **[WebNavigationBarSet](Publisher.WebNavigationBarSet.md)** object. Multiple pages in a web publication can each have a shape representing an instance of the same web navigation bar set. Changes made to a **WebNavigationBarSet** object are reflected in all the shapes representing instances of that web navigation bar set.


## Example

The following example tests to determine which shapes on the first page of the active document represent instances of web navigation bars. For each such shape found, the web navigation bar that it represents an instance of is set to auto update (see also the **[PbShapeType](publisher.pbshapetype.md)** enumeration).

```vb
Sub SetWebBarsToAutoUpdate() 
 
Dim shpLoop As Shape 
Dim strWebNavBarName As String 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.Type = pbWebNavigationBar Then 
 
 strWebNavBarName = shpLoop.WebNavigationBarSetName 
 With ActiveDocument.WebNavigationBarSets(strWebNavBarName) 
 .AutoUpdate = True 
 End With 
 
 End If 
Next 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]