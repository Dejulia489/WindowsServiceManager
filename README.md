# WindowsServiceManager

## Tasks Included

Windows Service Manager - Publish a windows service with a deployment group target.

1. Locates the service by the **Service Name**
2. Stops the service
3. If **Clean Install** is selected, removes all files in the parent directory of the .exe file
4. Starts the service

## Install and Run

After installing the Azure DevOps Pipelines extension from the link below, open a release and add the 'Windows Service Manager' task.
[Windows Service Manager](https://marketplace.visualstudio.com/items?itemName=MDSolutions.WindowsServiceManagerWindowsServiceManager)

### Task Options

![Task Options](https://github.com/Dejulia489/WindowsServiceManager/blob/master/Images/TaskOptions.png?raw=true "Task Options")

#### Required Fields

1. **Service Name** - The name of the Windows Service installed on the Deployment Group Target.
2. **Artifact Path** - The path to the Artifact that contains the Windows Service binaries.

#### Optional Fields

1. **Clean Install** - Will remove all files inside the parent direcotry of the .exe file prior to copying the Artifact.