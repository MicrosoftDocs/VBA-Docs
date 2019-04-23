---
title: Category.CategoryID property (Outlook)
keywords: vbaol11.chm2429
f1_keywords:
- vbaol11.chm2429
ms.prod: outlook
api_name:
- Outlook.Category.CategoryID
ms.assetid: e75ed17a-940f-2325-8739-1367329854d2
ms.date: 06/08/2017
localization_priority: Normal
---


# Category.CategoryID property (Outlook)

Returns a  **String** value that represents the unique identifier for the **[Category](Outlook.Category.md)** object. Read-only.


## Syntax

_expression_. `CategoryID`

_expression_ A variable that represents a [Category](Outlook.Category.md) object.


## Remarks

Because the  **[Name](Outlook.Category.Name.md)** property of a **Category** object can be changed either programmatically or by user action, each **Category** object is uniquely identified by a globally unique identifier (GUID), assigned to the object, that can be retrieved using this property. The GUID is presented as a string using the following format:


```vb
{00000000-0000-0000-0000-000000000000}
```


## Example

The following Visual Basic for Applications (VBA) example displays a dialog box containing the names and identifiers for each  **Category** object contained in the **[Categories](Outlook.NameSpace.Categories.md)** collection associated with the default **[NameSpace](Outlook.NameSpace.md)** object.


```vb
Private Sub ListCategoryIDs() 
 
 Dim objNameSpace As NameSpace 
 
 Dim objCategory As Category 
 
 Dim strOutput As String 
 
 
 
 ' Obtain a NameSpace object reference. 
 
 Set objNameSpace = Application.GetNamespace("MAPI") 
 
 
 
 ' Check if the Categories collection for the Namespace 
 
 ' contains one or more Category objects. 
 
 If objNameSpace.Categories.Count > 0 Then 
 
 
 
 ' Enumerate the Categories collection. 
 
 For Each objCategory In objNameSpace.Categories 
 
 
 
 ' Add the name and ID of the Category object to 
 
 ' the output string. 
 
 strOutput = strOutput & objCategory.Name & _ 
 
 ": " & objCategory.CategoryID & vbCrLf 
 
 Next 
 
 End If 
 
 
 
 ' Display the output string. 
 
 MsgBox strOutput 
 
 
 
 ' Clean up. 
 
 Set objCategory = Nothing 
 
 Set objNameSpace = Nothing 
 
 
 
End Sub
```


## See also


[Category Object](Outlook.Category.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]