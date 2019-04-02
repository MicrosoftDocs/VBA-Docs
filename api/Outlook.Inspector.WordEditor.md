---
title: Inspector.WordEditor property (Outlook)
keywords: vbaol11.chm2972
f1_keywords:
- vbaol11.chm2972
ms.prod: outlook
api_name:
- Outlook.Inspector.WordEditor
ms.assetid: 9e09b772-f679-19e6-905e-552ccadb0d24
ms.date: 06/08/2017
localization_priority: Normal
---


# Inspector.WordEditor property (Outlook)

Returns the Microsoft Word Document Object Model of the message being displayed. Read-only.


## Syntax

_expression_. `WordEditor`

_expression_ A variable that represents an [Inspector](Outlook.Inspector.md) object.


## Remarks

The  **WordEditor** property is only valid if the **[IsWordMail](Outlook.Inspector.IsWordMail.md)** method returns **True** and the **[EditorType](Outlook.Inspector.EditorType.md)** property is **olEditorWord**. The returned **WordDocument** object provides access to most of the Word object model except for the following members:


-  **Tables.Add**
    
-  **Range.ConvertToTable**
    
-  **InlineShapes.AddChart**
    
-  **Shapes.AddChart**
    
-  **Range.InsertXML**
    
-  **Selection.InsertXML**
    
-  **Range.ImportFragment**
    



## See also


[Inspector Object](Outlook.Inspector.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
