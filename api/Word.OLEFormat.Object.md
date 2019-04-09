---
title: OLEFormat.Object property (Word)
keywords: vbawd10.chm154337294
f1_keywords:
- vbawd10.chm154337294
ms.prod: word
api_name:
- Word.OLEFormat.Object
ms.assetid: 6f6a1c22-487a-d125-a759-43e9d659eaba
ms.date: 06/08/2017
localization_priority: Normal
---


# OLEFormat.Object property (Word)

Returns an  **Object** that represents the specified OLE object's top-level interface. .


## Syntax

_expression_.**Object**

 _expression_ An expression that returns an '[OLEFormat](Word.OLEFormat.md)' object.


## Remarks

This property allows you to access the properties and methods of an ActiveX control or the application in which an OLE object was created. The OLE object must support OLE Automation for this property to work.


## Example

This example sets the value of the first shape on the active document. For the example to work, this first shape must be an ActiveX control (for example, a check box or an option button).


```vb
With ActiveDocument.Shapes(1).OLEFormat 
 .Activate 
 Set myObj = .Object 
End With 
myObj.Value = True
```

This example adds a new ActiveX control to the active document. The example then activates the new option button and sets some of its properties.




```vb
Set myOB = ActiveDocument.Shapes _ 
 .AddOLEControl(ClassType:="Forms.OptionButton.1") 
With myOB.OLEFormat 
 .Activate 
 Set myObj = .Object 
End With 
With myObj 
 .Value = False 
 .Caption = "My Caption" 
 .AutoSize = True 
End With
```


## See also


[OLEFormat Object](Word.OLEFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]