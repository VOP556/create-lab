# create-lab
Create a hyper-v environment with powershell and WMF 5
The idea is to preprocess the physical and virtual machines in excel and then deploy them via powershell

in the end you will have a environment with a hyper-v cluster in the service area Infrastructrue (INFRA) and a virtualized Service Area (Plattform) and
then the same in the DMZ as a service delivery point to publish your services externaly

All will be done with focus on Objects and Powershell DSC 
In example the Node Objects are able to return their ConfigurationData to give it to a Powershell Desired State Configuration as Parameter.
