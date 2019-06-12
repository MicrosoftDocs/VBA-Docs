---
title: Plate.Index property (Publisher)
keywords: vbapb10.chm2883589
f1_keywords:
- vbapb10.chm2883589
ms.prod: publisher
api_name:
- Publisher.Plate.Index
ms.assetid: 7a16bd86-f0c4-d2df-832e-e9a55fed9068
ms.date: 06/13/2019
localization_priority: Normal
---


# Plate.Index property (Publisher)

Returns a **Long** that represents the position of a particular item in a specified collection. 


## Syntax

_expression_.**Index**

_expression_ A variable that represents a **[Plate](Publisher.Plate.md)** object.


## Example

The following example loops through the **MailMergeDataFields** collection and displays the **Index** and **Name** properties for each field.

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

The following example loops through the **Plates** collection and displays the **Index** and **Name** properties for each plate.

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