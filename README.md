# Office365_Service_Comms_2SCP
 Office 365 Messages via Service Communications API - To ASCII Text File - Send via SCP


    Instructions:
        Create AzureAD App with 'Office 365 Management APis' Permissions:
                        ServiceHealth.Read / Type: Application / Admin Consent Required: Yes

            Call with splats or 
        Edit $APIauthSettings hashtable with AzureAD App information
        Edit $SCPauthSettings hashtable with target SCP host information
