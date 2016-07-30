# vsix-deploy

An extension to allow simple running of a custom MsBuild task directly from Visual Studio.

Add the .proj-file (and import it in your projects .*proj-file) using __Tools__ -> __Create Deploy Task__

Then you will be able to run the task specified in ```Deploy.proj``` by simply right clicking the project and selecting __Deploy__, or by selecting __Build__ -> __Deploy PROJECTNAME__.
