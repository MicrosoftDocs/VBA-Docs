---
title: Shapes.AddOLEObject method (Publisher)
keywords: vbapb10.chm2162709
f1_keywords:
- vbapb10.chm2162709
ms.prod: publisher
api_name:
- Publisher.Shapes.AddOLEObject
ms.assetid: c454f9cb-2005-5e55-80a7-6dfbe9c109e5
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddOLEObject method (Publisher)

Adds a new **[Shape](Publisher.Shape.md)** object representing an OLE object to the specified **Shapes** collection.


## Syntax

_expression_.**AddOLEObject** (_Left_, _Top_, _Width_, _Height_, _ClassName_, _FileName_, _Link_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Left_|Required| **Variant**|The position of the left edge of the shape representing the OLE object.|
|_Top_|Required| **Variant**|The position of the top edge of the shape representing the OLE object.|
|_Width_|Optional| **Variant**|The width of the shape representing the OLE object. Default is -1, meaning that the width of the shape is automatically set based on the object's data.|
|_Height_|Optional| **Variant**|The height of the shape representing the OLE object. Default is -1, meaning that the height of the shape is automatically set based on the object's data.|
|_ClassName_|Optional| **String**|The class name of the OLE object to be added.|
|_FileName_|Optional| **String**|The file name of the OLE object to be added. If the path is not specified, the current working folder is used.|
|_Link_|Optional| **[MsoTriState](office.msotristate.md)**|Determines whether the OLE object is linked to or embedded in the publication.|

## Return value

Shape


## Remarks

For the _Left_, _Top_, _Width_, and _Height_ arguments, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

You must specify either a _ClassName_ or a _FileName_. If neither argument is specified, or if both are specified, an error occurs.

The _Link_ parameter can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|The OLE object is embedded.|
| **msoTrue**|The OLE object is linked. The default.|

## Example

The following example adds a Microsoft Office Excel worksheet to the first page of the active publication and activates the worksheet for editing.

```vb
Dim shpSheet As Shape 
 
Set shpSheet = ActiveDocument.Pages(1).Shapes.AddOLEObject _ 
 (Left:=72, Top:=72, ClassName:="Excel.Sheet") 
 
shpSheet.OLEFormat.Activate
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]