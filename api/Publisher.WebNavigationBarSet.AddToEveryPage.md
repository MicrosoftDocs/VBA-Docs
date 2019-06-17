---
title: WebNavigationBarSet.AddToEveryPage method (Publisher)
keywords: vbapb10.chm8519698
f1_keywords:
- vbapb10.chm8519698
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.AddToEveryPage
ms.assetid: d36a3281-a313-084c-0ae9-7a981a7d9713
ms.date: 06/18/2019
localization_priority: Normal
---


# WebNavigationBarSet.AddToEveryPage method (Publisher)

Adds a **[ShapeRange](publisher.shaperange.md)** object of type **pbWebNavigationBar** (**[PbShapeType](publisher.pbshapetype.md)** enumeration) to each page of the current document.


## Syntax

_expression_.**AddToEveryPage** (_Left_, _Top_, _Width_)

_expression_ A variable that represents a **[WebNavigationBarSet](Publisher.WebNavigationBarSet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Left_|Required| **Variant**|The position of the left edge of the shape representing the web navigation bar set.|
|_Top_|Required| **Variant**|The position of the top edge of the shape representing the web navigation bar set.|
|_Width_|Optional| **Variant**|The width of the shape representing the web navigation bar set.|

## Return value

ShapeRange


## Remarks

The specified web navigation bar set must exist before calling this method. 


## Example

The following example adds a web navigation bar set named WebNavBarSet1 to the top of every page in the active document.

```vb
ActiveDocument.WebNavigationBarSets("WebNavBarSet1") _ 
 .AddToEveryPage Left:=10, Top:=20 

```

<br/>

The following example adds a new web navigation bar set to the active document and adds it to every page of the publication.

```vb
Dim objWebNavBarSet As WebNavigationBarSet 
 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets.AddSet( _ 
 Name:="WebNavBarSet1", _ 
 Design:=pbnbDesignTopLine, _ 
 AutoUpdate:=True) 
 
objWebNavBarSet.AddToEveryPage Left:=50, Top:=10, Width:=500
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]