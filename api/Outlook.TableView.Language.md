---
title: TableView.Language property (Outlook)
keywords: vbaol11.chm2508
f1_keywords:
- vbaol11.chm2508
ms.prod: outlook
api_name:
- Outlook.TableView.Language
ms.assetid: cd600b12-0858-3edb-9c3a-5dc4cd0fc8bc
ms.date: 06/08/2017
localization_priority: Normal
---


# TableView.Language property (Outlook)

Returns or sets a **String** value that represents the language setting for the view. Read/write.


## Syntax

_expression_.**Language**

_expression_ A variable that represents a [TableView](Outlook.TableView.md) object.


## Remarks

The **Language** property uses a **String** to represent an ISO language tag. For example, the string "EN-US" represents the ISO code for "United States - English."

If a valid language code is specified, the object will only be available in the  **View** menu for the specified language type. If no value is specified, the object item is available for all language types. The default value for this property is an empty string.


## Example

The following Microsoft Visual Basic for Applications (VBA) example sets the language type of all  **[View](Outlook.View.md)** objects of type **olTableView** to U.S. English.


```vb
Sub SetLanguage() 
 
 'Sets the language of all table views to U.S. English. 
 
 Dim objViews As Outlook.Views 
 
 Dim objView As Outlook.View 
 
 
 
 Set objViews = _ 
 
 Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Views 
 
 'Iterate through each view in the collection. 
 
 For Each objView In objViews 
 
 Debug.Print objView.Name 
 
 'If view is of type olTableVIew then set language. 
 
 If objView.ViewType = olTableView And objView.Standard = False Then 
 
 objView.Language = "EN-US" 
 
 End If 
 
 Next objView 
 
End Sub
```


## See also


[TableView Object](Outlook.TableView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]