---
title: CustomProperty object (Word)
keywords: vbawd10.chm3552
f1_keywords:
- vbawd10.chm3552
ms.prod: word
api_name:
- Word.CustomProperty
ms.assetid: 1c4aa1ba-ad56-54d1-6e0d-2a82f7b9f4a9
ms.date: 06/08/2017
localization_priority: Normal
---


# CustomProperty object (Word)

Represents a single instance of a custom property for a smart tag. The **CustomProperty** object is a member of the **[CustomProperties](Word.CustomProperties.md)** collection.


## Remarks

Use the **[Item](Word.CustomProperties.Item.md)** method�or **[Properties](overview/Word.md)** (Index), where Index is the number of the property�of the **CustomProperties** collection to return a **CustomProperty** object.

Use the **[Name](Word.CustomProperty.Name.md)** and **[Value](Word.CustomProperty.Value.md)** properties to return the information related to a custom property for a smart tag. This example displays a message containing the name and value of the first custom property of the first smart tag in the current document. This example assumes that the current document contains at least one smart tag and that the first smart tag has at least one custom property.




```vb
Sub SmartTagsProps() 
 With ActiveDocument.SmartTags(Index:=1).Properties.Item(Index:=1) 
 MsgBox "Smart Tag Name: " & .Name & vbLf & _ 
 "Smart Tag Value: " & .Value 
 End With 
End Sub
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]