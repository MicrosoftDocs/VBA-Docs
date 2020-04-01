---
title: Send a Sharing Invitation for an RSS Feed
ms.prod: outlook
ms.assetid: 0b5b8ff5-d990-d869-7f80-15bbdcbec5a2
ms.date: 06/08/2017
localization_priority: Normal
---


# Send a Sharing Invitation for an RSS Feed

Sharing messages, including sharing invitations, sharing requests, and sharing responses, are represented in Microsoft Outlook by the **SharingItem](../../../api/Outlook.SharingItem.md)** object. The **[CreateSharingItem](../../../api/Outlook.NameSpace.CreateSharingItem.md)** method of the **[NameSpace](../../../api/Outlook.NameSpace.md)** object is used to create **SharingItem** objects for sharing invitations and sharing requests. Sharing responses are automatically created by Outlook when the **[Reply](../../../api/Outlook.SharingItem.Reply(method).md)** or **[ReplyAll](../../../api/Outlook.SharingItem.ReplyAll(method).md)** methods of a **SharingItem** that represents a sharing invitation or sharing request are called.

This sample uses the **penSharingItem** method to create a **SharingItem** that represents a sharing invitation for a Really Simple Syndication (RSS) feed. Once shared, the recipient can then use the **[OpenSharedFolder](../../../api/Outlook.NameSpace.OpenSharedFolder.md)** method of the **NameSpace** object or the **[OpenSharedFolder](../../../api/Outlook.SharingItem.OpenSharedFolder.md)** method of the **SharingItem** object to open the RSS feed.

1. The sample first creates a **ameSpace** object reference to the MAPI namespace.
    
2. It then uses the **reateSharingItem** method to create a new **SharingItem** object, using the URI of the RSS feed to establish the sharing context used by the **SharingItem**.
    
3. Finally, the **Add](../../../api/Outlook.Recipients.Add.md)** method for the **[Recipients](../../../api/Outlook.MailItem.Recipients.md)** collection of the newly created **SharingItem** object is called to add the specified recipient and the **[Send](../../../api/Outlook.SharingItem.Send(method).md)** method is used to send the **SharingItem**.

```vb
Public Sub ShareRSSByInvitation() 
 Dim oNamespace As NameSpace 
 Dim sRSSurl As String 
 Dim oSharingItem As SharingItem 
 
 On Error GoTo ErrRoutine 
 
 ' Specify the RSS feed URL for which sharing is to 
 ' be requested. 
 sRSSurl = "feed://example.com/rss.xml" 
 
 ' Get a reference to the MAPI namespace. 
 Set oNamespace = Application.GetNamespace("MAPI") 
 
 ' Create a new sharing request, using the RSS feed 
 ' URL to establish sharing context. 
 Set oSharingItem = oNamespace.CreateSharingItem(sRSSurl) 
 
 ' Add a recipient to the Recipients collection of 
 ' the sharing invitation. 
 oSharingItem.Recipients.Add "someone@example.com" 
 
 ' Send the sharing invitation. 
 oSharingItem.Send 
 
EndRoutine: 
 On Error GoTo 0 
 Set oSharingItem = Nothing 
 Set oFolder = Nothing 
 Set oNamespace = Nothing 
Exit Sub 
 
ErrRoutine: 
 Select Case Err.Number 
 Case 287 ' &H0000011F 
 ' The user denied access to the Address Book. 
 ' This error occurs if the code is run by an 
 ' untrusted application, and the user chose not to 
 ' allow access. 
 MsgBox "Access to Outlook was denied by the user.", _ 
 vbOKOnly, _ 
 Err.Number & " - " & Err.Source 
 Case -313393143 ' &HED520009 
 ' This error typically occurs if you set the 
 ' AllowWriteAccess property to true for a 
 ' default folder. 
 MsgBox Err.Description, _ 
 vbOKOnly, _ 
 Err.Number & " - " & Err.Source 
 Case -2147467259 ' &H80004005 
 ' This error typically occurs if the SharingItem 
 ' cannot be sent because of incorrect or 
 ' conflicting property settings. 
 MsgBox Err.Description, _ 
 vbOKOnly, _ 
 Err.Number & " - " & Err.Source 
 Case Else 
 ' Any other error that may occur. 
 MsgBox Err.Description, _ 
 vbOKOnly, _ 
 Err.Number & " - " & Err.Source 
 End Select 
 
 GoTo EndRoutine 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]