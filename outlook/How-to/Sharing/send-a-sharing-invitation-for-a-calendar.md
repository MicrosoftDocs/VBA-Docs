---
title: Send a Sharing Invitation for a Calendar
ms.prod: outlook
ms.assetid: 830f0c51-251c-f0f4-71b8-6090089022c5
ms.date: 06/08/2017
localization_priority: Normal
---


# Send a Sharing Invitation for a Calendar

Sharing messages, including sharing invitations, sharing requests, and sharing responses, are represented in Microsoft Outlook by the **[SharingItem](../../../api/Outlook.SharingItem.md)** object. The **[CreateSharingItem](../../../api/Outlook.SharingItem.Recipients.md)** method of the **[NameSpace](../../../api/Outlook.NameSpace.md)** object is used to create **SharingItem** objects for sharing invitations and sharing requests.

This sample uses the **OpenSharingItem** method to create a **SharingItem** that represents sharing invitation for your **Calendar** default folder. Once shared, the recipient can then use the **[OpenSharedFolder](../../../api/Outlook.NameSpace.OpenSharedFolder.md)** or **[GetSharedDefaultFolder](../../../api/Outlook.NameSpace.GetSharedDefaultFolder.md)** methods of the **NameSpace** object, or the **[OpenSharedFolder](../../../api/Outlook.SharingItem.OpenSharedFolder.md)** method of the **SharingItem** object to open the shared folder.

1. The sample obtains a **[Folder](../../../api/Outlook.Folder.md)** object reference for the **Calendar** default folder for the current user, by using the **[GetDefaultFolder](../../../api/Outlook.NameSpace.GetDefaultFolder.md)** method of the **NameSpace** object.
    
2. It uses the **CreateSharingItem** method to create a new **SharingItem** object, using the **Folder** object to establish the sharing context used by the **SharingItem**.
    
3. Finally, the **[Add](../../../api/Outlook.Recipients.Add.md)** method for the **[Recipients](../../../api/Outlook.MailItem.Recipients.md)** collection of the newly created **SharingItem** object is called to add the specified recipient and the **[Send](../../../api/Outlook.SharingItem.Send(method).md)** method is used to send the **SharingItem**.
    
```vb
Public Sub ShareCalendarByInvitation() 
 Dim oNamespace As NameSpace 
 Dim oFolder As Folder 
 Dim oSharingItem As SharingItem 
 
 On Error GoTo ErrRoutine 
 
 ' Get a reference to the Calendar default folder 
 Set oNamespace = Application.GetNamespace("MAPI") 
 Set oFolder = oNamespace.GetDefaultFolder(olFolderCalendar) 
 
 ' Create a new sharing invitation, using the Calendar 
 ' default folder to establish sharing context. 
 Set oSharingItem = oNamespace.CreateSharingItem(oFolder) 
 
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
 ' AllowWriteAccess property of a SharingItem 
 ' to True when sharing a default folder. 
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