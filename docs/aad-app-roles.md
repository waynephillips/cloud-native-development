
# Adding Custom Roles in your AAD Application
When implementing app level Authorization, it's possible to define Roles/Claims in AAD, instead of relying directly on AD Groups.

Defining Application Roles in Azure AD has the following benefits:
- Application users and user roles are visible/can be managed through the Azure Portal/B2B Invite API.
- You can assign AD Groups to your custom roles.
- Role information is provided in a user's id_token, reducing round-trips to MS Graph API to verify access.

You can read much more about this topic [here](https://www.microsoftpressstore.com/articles/article.aspx?p=2473127).


## Provisioning the Application Roles in new App Registrations

1. Assuming you have access to the App Registration, go to the  **Settings** blade for the your App Registraiton.
2. Click **Manifest** > **Edit**.
3. Add the following JSON to the `appRoles` element. [Generate new GUIDs](https://www.guidgenerator.com/online-guid-generator.aspx) for the `id` properties.

    ```json
    {
      "allowedMemberTypes": [
        "User"
      ],
      "displayName": "<The Display Name of your role>",
      "id": "<Generate a new GUID>",  
      "isEnabled": true,
      "description": "<Role Description>",
      "value": "<ValueOfRoleInToken>"
    },
    {
      "allowedMemberTypes": [
        "User"
      ],
      "displayName": "InvitationApprover",
      "id": "<Generate a new GUID>",  
      "isEnabled": true,
      "description": "Invitation approvers may approve invite requests.",
      "value": "InvitationApprover"
    }
    ```
