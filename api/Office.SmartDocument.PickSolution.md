---
title: SmartDocument.PickSolution method (Office)
keywords: vbaof11.chm262003
f1_keywords:
- vbaof11.chm262003
ms.prod: office
api_name:
- Office.SmartDocument.PickSolution
ms.assetid: ea50c7a4-4b52-10c4-8b1a-86c7ef80dec1
ms.date: 01/25/2019
ms.localizationpriority: medium
---


# SmartDocument.PickSolution method (Office)

Displays a dialog box that allows the user to choose an available XML expansion pack to attach to the active document in Microsoft Word or to a workbook in Microsoft Excel.


## Syntax

_expression_.**PickSolution** (_ConsiderAllSchemas_)

_expression_ A variable that represents a **[SmartDocument](Office.SmartDocument.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ConsiderAllSchemas_|Optional|**Boolean**|**True** displays all available XML expansion packs installed on the user's computer. **False** displays only XML expansion packs applicable to the active document. Default value is **False**.|

## Remarks

Use the **PickSolution** method to allow the user to select an XML expansion pack from a list. The schema attached to the active document or workbook determines which XML expansion packs are applicable.

The **PickSolution** method does not return a value to indicate whether the user selected an XML expansion pack or chose **Cancel** in the dialog box. Check the **SolutionID** property after calling **PickSolution** to determine whether an XML expansion pack has been attached.

If the smart document developer has failed to specify "targetApplication" in the XML expansion pack manifest file, the list displayed by **PickSolution** may include XML expansion packs that are not targeted to the active application; for example, an Excel user may see XML expansion packs targeted exclusively to Word. In these circumstances, the user may select an XML expansion pack that is not appropriate for the active application.

For more information about smart documents or XML expansion packs for smart documents, see the [Smart Document Software Development Kit (SDK)](/previous-versions/office/developer/office-2003/aa193924(v%3doffice.11)).


## Example

The following example checks the **SolutionID** property to determine whether the active Microsoft Word document already has an attached XML expansion pack; if not, it displays a dialog box that allows the user to choose an available XML expansion pack. It then displays the properties of the smart document.


```vb
 Dim objSmartDoc As Office.SmartDocument 
 Dim strSmartDocInfo As String 
 Set objSmartDoc = ActiveDocument.SmartDocument 
 If objSmartDoc.SolutionID = "None" Or objSmartDoc.SolutionID = "" Then 
 objSmartDoc.PickSolution True 
 End If 
 If objSmartDoc.SolutionID > "None" And objSmartDoc.SolutionID > "" Then 
 strSmartDocInfo = "SolutionID: " & objSmartDoc.SolutionID & vbCrLf & _ 
 "SolutionURL: " & objSmartDoc.SolutionURL 
 MsgBox strSmartDocInfo, vbInformation + vbOKOnly, "Smart Doc Properties" 
 Else 
 MsgBox "The user clicked Cancel." 
 End If 
 Set objSmartDoc = Nothing 
 

```


## See also

- [SmartDocument object members](overview/Library-Reference/smartdocument-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]