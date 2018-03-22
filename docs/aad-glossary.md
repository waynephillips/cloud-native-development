From: https://msdn.microsoft.com/Library/Azure/Ad/Graph/api/entity-and-complex-type-reference

# Entities
**AppRoles** - The collection of application roles that an application may declare. These roles can be assigned to users, groups or service principals.

**AppRoleAssignment** - Used to record when a user or group is assigned to an application. In this case, the role assignment will result in an application tile showing up on the user's app access panel. This entity may also be used to grant another application (modeled as a service principal) access to a resource application in a particular role. You can create, read, update, and delete role assignments. Inherits from DirectoryObject.

**OAuth2PermissionGrant** - Represents the OAuth 2.0 delegated permission scopes that have been granted to an application (represented by a service principal) as part of the user or admin consent process.

# Relationships
**User.AppRoleAssignments** - The set of applications that this user is assigned to.

**Group.AppRoleAssignments** - Contains the set of applications that a group is assigned to.

**ServicePrincipal.AppRoleAssignments** - Applications that the service principal is assigned to (application permissions).

**ServicePrincipal.AppRoleAssignedTo** - Principals (users, groups, and service principals) that are assigned to this service principal.

**User.Oauth2PermissionGrants** - The set of applications that are granted consent to impersonate this user (delegated permissions).

**ServicePrincipal.Oauth2PermissionGrants** - User impersonation grants associated with this service principal (delegated permissions).

**Application.AppRoles** - The collection of application roles that an application may declare. These roles can be assigned to users, groups or service principals.

# Others
**DirectoryRoles** - completely unrelated from applications and grant users permissions to perform operations on the directory.

More info on those here: https://azure.microsoft.com/en-us/documentation/articles/active-directory-assign-admin-roles/
