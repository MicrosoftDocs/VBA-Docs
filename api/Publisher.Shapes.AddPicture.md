---
title: Shapes.AddPicture method (Publisher)
keywords: vbapb10.chm2162710
f1_keywords:
- vbapb10.chm2162710
ms.prod: publisher
api_name:
- Publisher.Shapes.AddPicture
ms.assetid: a5305bd0-295f-46f6-7823-46dab750243b
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddPicture method (Publisher)

Adds a new **[Shape](Publisher.Shape.md)** object representing a picture to the specified **Shapes** collection.


## Syntax

_expression_.**AddPicture** (_FileName_, _LinkToFile_, _SaveWithDocument_, _Left_, _Top_, _Width_, _Height_)

_expression_ A variable that represents a **[Shapes](publisher.shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_FileName_|Required| **String**|The name of the picture file to insert into the shape. The path can be absolute or relative.|
|_LinkToFile_|Required| **[MsoTriState](office.msotristate.md)**|Determines whether the picture is linked to or embedded in the publication.|
|_SaveWithDocument_|Required| **MsoTriState**|Determines whether the picture is saved as a separate file with the publication.|
|_Left_|Required| **Variant**|The position of the left edge of the shape representing the picture.|
|_Top_|Required| **Variant**|The position of the top edge of the shape representing the picture.|
|_Width_|Optional| **Variant**|The width of the shape representing the picture. Default is -1, meaning that the width of the shape is automatically set based on the object's data.|
|_Height_|Optional| **Variant**|The height of the shape representing the picture. Default is -1, meaning that the height of the shape is automatically set based on the object's data.|

## Return value

Shape


## Remarks

If the _SaveWithDocument_ argument is **msoTrue**, Microsoft Publisher saves a new copy of the picture file specified by the _FileName_ argument in the same directory as the publication.

The _LinkToFile_ and _SaveWithDocument_ arguments cannot have the same value, or else an error occurs. If either argument is **msoTrue**, the other must be **msoFalse**.

For the _Left_, _Top_, _Width_, and _Height_ arguments, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Publisher (for example, "2.5 in").

The _LinkToFile_ parameter can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|The picture is to be embedded in the publication.|
| **msoTrue**|The picture is to be linked to the publication.|

## Example

The following example adds a picture based on an existing file to the active publication; the picture in the publication is linked to a copy of the original file. Note that `PathToFile` must be replaced with a valid file path for this example to work.

```vb
Dim shpPicture As Shape 
 
Set shpPicture = ActiveDocument.Pages(1).Shapes.AddPicture _ 
 (FileName:="PathToFile", _ 
 LinkToFile:=msoTrue, _ 
 SaveWithDocument:=msoFalse 
 Left:=72, Top:=72)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]