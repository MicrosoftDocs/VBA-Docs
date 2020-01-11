---
title: Identify a Meeting Item as a Counter-Proposal to a Prior Meeting Request
ms.prod: outlook
ms.assetid: 42d53f48-d9de-18d8-d39b-86feceff0eaa
ms.date: 06/08/2017
localization_priority: Normal
---


# Identify a Meeting Item as a Counter-Proposal to a Prior Meeting Request

This topic shows how to use the named property,  [PidLidAppointmentCounterProposal](../../../api/overview/Outlook.md), and the Microsoft Outlook object model to identify a  **[MeetingItem](../../../api/Outlook.MeetingItem.md)** object as a counter proposal to a prior meeting request.

In the Outlook object model, all types of items, such as a mail item and a contact item, correspond to specific message classes. In particular, responses to a meeting request can be identified by the following message classes: 

- IPM.Schedule.Meeting.Resp.Neg for a decline response
    
- IPM.Schedule.Meeting.Resp.Pos for an acceptance response
    
- IPM.Schedule.Meeting.Resp.Ten for a tentative response
    

However, the Outlook object model does not provide a means to identify a response as the fourth possible response to a meeting request, which is a counter-proposal.
Using the  **[PropertyAccessor](../../../api/Outlook.PropertyAccessor.md)** object and the **PSETID_Appointment** namespace definition of **PidLidAppointmentCounterProposal**, you can program within the object model to distinguish all responses of a meeting request item. The following code sample in C# shows how to get the property value given a meeting item. Note that in the code sample, the named property is expressed as: 



```vb
"https://schemas.microsoft.com/mapi/id/00062002-0000-0000-C000-000000000046}/8257000B"
```

where  `{00062002-0000-0000-C000-000000000046}` is the **PSETID_Appointment** namespace and `8257000B` is the property tag of **PidLidAppointmentCounterProposal**.



```cs
private bool IsCounterProposal(Outlook.MeetingItem meeting) 
{ 
    const string counterPropose = 
        "https://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8257000B"; 
    Outlook.PropertyAccessor pa = meeting.PropertyAccessor; 
    if ((bool)pa.GetProperty(counterPropose)) 
        return true; 
    else 
        return false;  
}
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]