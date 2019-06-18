---
title: WebNavigationBarSets.AddSet method (Publisher)
keywords: vbapb10.chm8454148
f1_keywords:
- vbapb10.chm8454148
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSets.AddSet
ms.assetid: 5b998e14-b1eb-2a4a-2ed5-9a1ef16d69c1
ms.date: 06/18/2019
localization_priority: Normal
---


# WebNavigationBarSets.AddSet method (Publisher)

Adds a new **[WebNavigationBarSet](Publisher.WebNavigationBarSet.md)** object representing a web navigation bar set to the specified **WebNavigationBarSets** collection. 


## Syntax

_expression_.**AddSet** (_Name_, _Design_, _AutoUpdate_)

_expression_ A variable that represents a **[WebNavigationBarSets](Publisher.WebNavigationBarSets.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Name_|Required| **String**|The name of the web navigation bar to be added. This parameter must be unique.|
|_Design_|Optional| **[PbWizardNavBarDesign](publisher.pbwizardnavbardesign.md)**|Specifies the navigation bar design scheme.|
|_AutoUpdate_|Optional| **Boolean**| **True** if all pages with the _AddHyperlinkToWebNavBar_ parameter (**[Pages.Add](publisher.pages.add.md)** method) set to **True** are added as links to the navigation bar, and the navigation bar is kept updated.|

## Return value

WebNavigationBarSet


## Remarks

The _Name_ parameter must be unique to avoid a run-time error.


## Example

The following example adds a **WebNavigationBarSet** object to the **WebNavigationBarSets** collection of the active document, and then sets some properties.

```vb
Dim objWebNavBarSet As WebNavigationBarSet 
 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets.AddSet( _ 
 Name:="WebNavBarSet1", _ 
 Design:=pbnbDesignAmbient, _ 
 AutoUpdate:=True) 
 
With objWebNavBarSet 
 .AddToEveryPage Left:=50, Top:=10 
 .ButtonStyle = pbnbDesignTopLine 
 .ChangeOrientation pbNavBarOrientHorizontal 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]