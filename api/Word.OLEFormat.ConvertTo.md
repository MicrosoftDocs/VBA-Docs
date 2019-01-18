---
title: OLEFormat.ConvertTo method (Word VBA)
keywords: vbawd10.chm154337390
f1_keywords:
- vbawd10.chm154337390
ms.prod: word
api_name:
- Word.OLEFormat.ConvertTo
ms.assetid: 6d648f38-34fa-21b1-3ab9-a1965f2398f4
ms.date: 12/05/2018
localization_priority: Normal
---


# OLEFormat.ConvertTo method (Word)

Converts the specified OLE object from one class to another, making it possible for you to edit the object in a different server application or change how the object is displayed in the document.


## Syntax

_expression_.**ConvertTo** ( _ClassType_, _DisplayAsIcon_, _IconFileName_, _IconIndex_, _IconLabel_ )

_expression_ Required. A variable that represents an [OLEFormat](Word.OLEFormat.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ClassType_|Optional| **Variant**|The name of the application used to activate the OLE object. You can see a list of the available applications in the **Object type** box on the **Create New** tab in the **Object** dialog box. You can find the ClassType string by inserting an object as an inline shape and then viewing the field codes. The class type of the object follows either the word "EMBED" or the word "LINK."|
| _DisplayAsIcon_|Optional| **Variant**| **True** to display the OLE object as an icon. The default value is **False**.|
| _IconFileName_|Optional| **Variant**|The file that contains the icon to be displayed.|
| _IconIndex_|Optional| **Variant**|The index number of the icon within _IconFileName_. The order of icons in the specified file corresponds to the order in which the icons appear in the **Change Icon** dialog box (**Insert Object** dialog box) when the **Display as icon** check box is selected.<br/><br/>The first icon in the file has the index number 0 (zero). If an icon with the given index number doesn't exist in _IconFileName_, the icon with the index number 1 (the second icon in the file) is used. The default value is 0 (zero).|
| _IconLabel_|Optional| **Variant**|A label (caption) to be displayed beneath the icon.|

## Example

This example creates a new document, and then inserts an embedded Word document with some text. The embedded document is then converted to a Word Picture.


```vb
Dim objEmbedded As Object 
 
Documents.Add 
 
Set objEmbedded = ActiveDocument.Shapes _ 
  .AddOLEObject(ClassType:= "Word.Document") 
objEmbedded.OLEFormat.Activate 
Selection.TypeText "Test" 
objEmbedded.OLEFormat.ConvertTo _ 
  ClassType:="Word.Picture"

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]