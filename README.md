# WindowsServiceManager

## Tasks Included

Windows Service Manager - Installs and deploys a windows service or [TopShelf](https://github.com/Topshelf/Topshelf) service on a target machine or a deployment group target.

1. Locates the service by the **Service Name**
2. Stops the service
3. If **Stop Process** is selected and the **Timeout** is reached, the process of the service will be stopped.
4. If **Clean Install** is selected, removes all files in the parent directory of the .exe file
5. Starts the service

## Install and Run

After installing the Azure DevOps Pipelines extension from the link below, open a release and add the 'Windows Service Manager' task.
[Windows Service Manager](https://marketplace.visualstudio.com/items?itemName=MDSolutions.WindowsServiceManagerWindowsServiceManager)

### Task Options for Agent

![Task Options](https://github.com/Dejulia489/WindowsServiceManager/blob/master/Images/TaskOptionsAgent.png?raw=true "Task Options Agent")

### Task Options for Deployment Group

![Task Options](https://github.com/Dejulia489/WindowsServiceManager/blob/master/Images/TaskOptionsDeploymentGroup.png?raw=true "Task Options Deployment Group")

### Advanced Task Options

![Task Options](https://github.com/Dejulia489/WindowsServiceManager/blob/master/Images/TaskOptionsAdvanced.png?raw=true "Advanced Task Options")

### Installation Task Options

![Task Options](https://github.com/Dejulia489/WindowsServiceManager/blob/master/Images/TaskOptionsInstallation.png?raw=true "Installation Task Options")

#### Required Fields

1. **Deployment Type** - Deploying with either an Agent or a Deployment Group Target.
2. **Service Name** - The name of the Windows Service installed on the Deployment Group Target.
3. **Artifact Path** - The path to the Artifact that contains the Windows Service binaries.
4. **Timeout** - The number of seconds to wait for the service to stop.

#### Required Fields with Agent Deployment Type

1. **Machines** - Provide a comma separated list of machine IP addresses or FQDNs. Or provide output variable of other tasks. Eg: $(variableName).
2. **Admin Login** - Administrator login for the target machines.
3. **Password** - Password for administrator login for the target machines. It can accept variable defined in Build/Release definitions as '$(passwordVariable)'. You may mark variable type as 'secret' to secure it.

#### Advanced Fields

1. **Stop Process** - Stops the process if the service does not respond within the timeout.
2. **Clean Install** - Removes all files inside the parent directory of the .exe file prior to copying the Artifact.

#### Installation Fields

1. **Install the windows service** - Enables service installation.
2. **Installation Path** - The path to the executable file.
3. **Run As Username** - The username the service should run as.
4. **Run As Password** - The password for the Run As Username. It can accept variable defined in Build/Release definitions as '$(passwordVariable)'. You may mark variable type as 'secret' to secure it.
5. **Install as a TopShelf Service** - Enables [TopShelf](https://github.com/Topshelf/Topshelf) installation.
6. **Instance Name** - The name of the [TopShelf](https://github.com/Topshelf/Topshelf) instance.
7. **Install Arguments** - The TopShelf installation arguments.

## Release Notes

[![Build status](https://dev.azure.com/MaDSolutionsLLC/Windows Service Manager/_apis/build/status/WindowsServiceManager)](https://dev.azure.com/MaDSolutionsLLC/Windows Service Manager/_build/latest?definitionId=3)

### Version 4

Implemented support for installing [TopShelf](https://github.com/Topshelf/Topshelf) windows services.

### Version 3

Implemented feature request for installing a service and resolved bug for multiple machine support.

### Version 2

Resolved bug for service name matching, service name will only support exact names.

### Version 1

Supports deployment groups and agents using WinRM.

### Version 0

Supports deployment groups only.