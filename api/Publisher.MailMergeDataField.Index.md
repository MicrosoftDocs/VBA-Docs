---
title: MailMergeDataField.Index property (Publisher)
keywords: vbapb10.chm6422529
f1_keywords:
- vbapb10.chm6422529
ms.prod: publisher
api_name:
- Publisher.MailMergeDataField.Index
ms.assetid: f70d0266-0527-6871-632d-b45b617d75d4
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataField.Index property (Publisher)

Returns a **Long** that represents the position of a particular item in a specified collection. 


## Syntax

_expression_.**Index**

_expression_ A variable that represents a **[MailMergeDataField](Publisher.MailMergeDataField.md)** object.


## Example

The following example loops through the **[MailMergeDataFields](Publisher.MailMergeDataFields.md)** collection and displays the **Index** and **Name** properties for each field.

```vb
Dim mmfLoop As MailMergeDataField 
 
With ActiveDocument.MailMerge.DataSource 
 If .DataFields.Count > 0 Then 
 For Each mmfLoop In .DataFields 
 Debug.Print "Field " & mmfLoop.Name _ 
 & " / Index " & mmfLoop.Index 
 Next mmfLoop 
 Else 
 Debug.Print "No fields to report." 
 End If 
End With
```

<br/>

The following example loops through the **[Plates](publisher.plates.md)** collection and displays the **Index** and **Name** properties for each plate.

```vb
Dim plaLoop As Plate 
 
If ActiveDocument.Plates.Count > 0 Then 
 For Each plaLoop In ActiveDocument.Plates 
 Debug.Print "Plate " & plaLoop.Name _ 
 & " / Index " & plaLoop.Index 
 Next plaLoop 
Else 
 Debug.Print "No plates to report." 
End If
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]