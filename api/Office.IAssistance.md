---
title: IAssistance Object (Office)
ms.prod: office
api_name:
- Office.IAssistance
ms.assetid: c8327d45-a6a2-dc4c-67f0-d02598eb60ba
ms.date: 06/08/2017
---


# IAssistance Object (Office)

Provides a means for developers to create a customized help experience for users within Microsoft Office.


## Remarks

The  **Assistance** property returns an **IAssistance** object. The **IAssistance** object exposes methods that allow developers to display help topics in the Office Help Viewer or to display help topics that ship with Office in the Help window of the host application. Developers either pass specific Help IDs to the help system or pass specific search queries. Help IDs have to be explicitly added to the Help file in order for the Help ID to return the help topic.


## Example

In the first line in the following example, the  **ShowHelp** method of the **IAssistance** object displays the help topic associated with the help ID "xlmain11.chm60407" in the "Excel" namespace. When the second line is executed, the client viewer will remain open and display the help topic associated with the help ID "65879", in the "Excel Developer" namespace.


```vb
Sub DisplayHelpTopic() 
 Application.Assistance.ShowHelp "xlmain11.chm60407", "" 
 Application.Assistance.ShowHelp "vbaxl10.chm65879", "DEV" 
End Sub 

```


## Methods



|**Name**|
|:-----|
|[ClearDefaultContext](Office.IAssistance.ClearDefaultContext.md)|
|[SearchHelp](Office.IAssistance.SearchHelp.md)|
|[SetDefaultContext](Office.IAssistance.SetDefaultContext.md)|
|[ShowHelp](Office.IAssistance.ShowHelp.md)|

## See also





[Object Model Reference](./overview/reference-object-library-reference-for-office.md)
