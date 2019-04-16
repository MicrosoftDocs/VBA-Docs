---
title: BusinessCardView.Copy method (Outlook)
keywords: vbaol11.chm2922
f1_keywords:
- vbaol11.chm2922
ms.prod: outlook
api_name:
- Outlook.BusinessCardView.Copy
ms.assetid: 9a0a1a14-87bd-ff53-6643-5e11a07733a1
ms.date: 06/08/2017
localization_priority: Normal
---


# BusinessCardView.Copy method (Outlook)

Creates a new  **[View](Outlook.View.md)** object based on the existing **[BusinessCardView](Outlook.BusinessCardView.md)** object.


## Syntax

_expression_.**Copy** (_Name_, _SaveOption_)

 _expression_ An expression that returns a [BusinessCardView](Outlook.BusinessCardView.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new view.|
| _SaveOption_|Optional| **[OlViewSaveOption](Outlook.OlViewSaveOption.md)**|The save option for the new view.|

## Return value

A  **View** object that represents the new view.


## Example

The following Visual Basic for Applications (VBA) example creates a copy of a  **BusinessCardView** object, named "New Card View", and saves it in the **Contacts** default folder. To run this example, you need to first create a **BusinessCardView** object named "Card View" either programmatically or by using the Microsoft Outlook user interface.


```vb
Sub CopyBusinessCardView() 
 
 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objNewView As BusinessCardView 
 
 
 
 ' Get the Views collection of the Contacts default folder. 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderContacts).Views 
 
 
 
 ' Copy the existing view. 
 
 Set objNewView = objViews("Card View").Copy( _ 
 
 "New Card View", _ 
 
 olViewSaveOptionThisFolderEveryone) 
 
 
 
End Sub
```


## See also


[BusinessCardView Object](Outlook.BusinessCardView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]