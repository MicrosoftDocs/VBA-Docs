---
title: BoundObjectFrame.SizeMode Property (Access)
keywords: vbaac10.chm10911
f1_keywords:
- vbaac10.chm10911
ms.prod: access
api_name:
- Access.BoundObjectFrame.SizeMode
ms.assetid: 2c44b16f-cb04-8e45-2a67-7424342f48de
ms.date: 06/08/2017
---


# BoundObjectFrame.SizeMode Property (Access)

You can use the  **SizeMode** property to specify how to size a picture or other object in a bound object frame, an unbound object frame, or an image control.


## Syntax

 _expression_. **SizeMode**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

The  **SizeMode** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Clip|**acOLESizeClip**|(Default) Displays the object at actual size. If the object is larger than the control, its image is clipped on the right and bottom by the control's borders.|
|Stretch|**acOLESizeStretch**|Sizes the object to fill the control. This setting may distort the proportions of the object.|
|Zoom|**acOLESizeZoom**|Displays the entire object, resizing it as necessary without distorting the proportions of the object. This setting may leave extra space in the control if the control is resized.|
Use the Clip setting for the fastest display. You can use the Stretch setting for bar graphs and line graphs without concern for size adjustments. The Stretch setting can distort circles and photos.


## Example

The following example creates a linked OLE object using an unbound object frame named  `OLE1` and sizes the control to display the object's entire contents when the user clicks a command button.


```vb
Sub Command1_Click 
 OLE1.Class = "Excel.Sheet" ' Set class name. 
 ' Specify type of object. 
 OLE1.OLETypeAllowed = acOLELinked 
 ' Specify source file. 
 OLE1.SourceDoc = "C:\Excel\Oletext.xls" 
 ' Specify data to create link to. 
 OLE1.SourceItem = "R1C1:R5C5" 
 ' Create linked object. 
 OLE1.Action = acOLECreateLink 
 ' Adjust control size. 
 OLE1.SizeMode = acOLESizeZoom 
End Sub
```


## See also


#### Concepts


[BoundObjectFrame Object](Access.BoundObjectFrame.md)

