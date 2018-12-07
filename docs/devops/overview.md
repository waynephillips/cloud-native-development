# DevOps Overview

## Introduction
This page provides a high level overview of the general project software development workflow in place for the Chicago DMC. In addition, it describes the DevOps-related processes and tools used for software development.

## Roles
The following roles are discussed in this document:
- AS – 
- Operations team (Ops) – Operations team
- Development team (Dev) – Developer team
- Embedded security - 

## Project Workflow
[[/images/path/to/software-dev-phases.png|alt=Software Development Phases]]
### Business requirements
- AS defines product requirements with business partner

### Project Planning
- Initial sprint and capacity planning
- Architectural design review

### Project Setup Phase
The project setup phase typically takes place in Sprint 0 and typically includes the following major steps
- Create project repository
- Provision and configure cloud resources using Infrastructure as Code (IaC)
- Set up build and release pipelines using Continuous Integration (CI) and Continuous Delivery (CD)
- Create cloud app registrations using IaC, request admin consent
- Request developer access to created cloud resources
- Request adding necessary Azure SQL database roles and users for automation and development
- Request adding user groups to app registration
- Initiate risk assessment

### Development Phase
The software development process is described in the following diagram:
 
In addition to the software development activities described above, the following ongoing processes happen during the Development Phase:
- Risk assessment (Typical duration: weeks)
- Refinements in infrastructure configurations using IaC

### Production Support
TBD

## DevOps

### Definitions

#### DevOps
DevOps is the union of people, process, and tools to enable continuous delivery of value to our end users. 

#### Continuous Integration
Continuous Integration (CI) is a development practice that requires developers to integrate code into a shared repository several times a day. Each check-in is then verified by an automated build, allowing teams to detect problems early.

#### Continuous Delivery
Continuous delivery is a series of practices designed to ensure that code can be rapidly and safely deployed to production by delivering every change to a production-like environment and ensuring business applications and services function as expected through automated testing. Since every change is delivered to a staging environment using complete automation, you can have confidence the application can be deployed to production with a push of a button when the business is ready.

### Application Delivery Pipeline
The following diagram describes the application delivery workflow:
 
In the Continuous Integration (CI) phase, a developer applies changes to the code base on a feature branch and initiates a pull request (PR). Code validation and testing is triggered automatically as part of the PR. Merging to the Master branch kicks off further validation and testing and eventually the publishing step. It is recommended that one create a PR early in the development process so that others can provide timely feedback on the implementation.

The Continuous Delivery (CD) phase is automatically triggered after the publishing step from CI is complete. The code package is first deployed to a Development (Dev) environment. The development team can send the code update further to the Test environment once the application state is ready for review by the AS and typically one or more business partners. Once the AS and the Ops team approve the changes present in Test, the changes are automatically deployed to production.

Note that a project repository includes not only the application code but also all necessary infrastructure configurations by leveraging Infrastructure as Code (IaC). More on this in the Infrastructure as Code section.

### Infrastructure as Code

#### Resource Library
In order to ensure reusability and consistency across environments and projects, Infrastructure as Code (IaC) is used for environment creation and modification. Furthermore, in order to simplify the provisioning of new cloud resources and provide default configurations, infrastructure templates are available in a centralized repository, called the Resource Library (RL). This a version control repository where IaC templates are stored. The templates are primarily developed and provided by the Ops team and go through rigorous infrastructure testing before making them available for use to development teams.
The following integration pipeline is used for publishing IaC templates:
 
#### Application Repository
Application projects often need custom modifications and updates in infrastructure configuration. In order to enable for this scenario, project-specific configurations are stored in the application repository in the form of scripts.

#### All Together: CI/CD and IaC
The following diagram illustrates the CI/CD workflow combined with IaC:
 
Note the two different version control sources. The app repository includes both the application code base and Configuration-as-Code (CaC) scripts. The list of IaC templates in use is defined in an Azure DevOps release pipeline along with any other deployment steps. The CD pipeline is the same as the one introduced in the Application Delivery Pipeline section.

To combine different infrastructure templates and thereby define custom infrastructure setups (cloud architectures), Azure DevOps is used as a tool. Leveraging Azure DevOps, one can quickly define the right combination of IaC templates.

#### More on Approval Steps
The following workflow describes the approval process in more detail:
 
Note the difference between post-approval performed by AS in the Test stage and the pre-approval done by the Ops team before deploying to production. 

### Tools in Use

#### Primary Technology stack
- Azure cloud (PaaS) – Cloud
- Azure Active Directory – Identity and Access Management
- .NET and .NET Core – Software framework
- Entity Framework and Entity Framework Core – ORM framework
#### DevOps Tools
- Azure DevOps – Delivery management and orchestration, CI/CD
- Checkmarx – Static code scan analysis
- Azure ARM templates – IaC
- Powershell – IaC and CaC

## Security
 
### Access