# Interop Test

This repo demonstrates a possible bug on Windows 7. It occurs when a .NET user control's event is handled by VB6 code, and that VB6 code instantiates a .NET class exposed to COM for the very first time in the program.

## Important Notes:
1. **We only get the issue in Windows 7 SP1** after updates have been installed. This works fine on Windows 10 and on a clean install of Windows 7 SP1. However, after running windows updates, we start getting the error in Windows 7.
2. **The error does not occur in development.** Registering the .NET dll with regasm makes it always work. VB6 only uses registered dlls as references so you cannot reproduce it in development (unless there is some magic you can do for reg-free assembly loading in the VB6 IDE that I don't know about.)
3. **Instantiating the class before the event makes it work.** The error only occurs if the class instantiated in the VB6 handler is created for the first time.

## The flows of the code is:

1. Start the application.
2. Create the form with the .NET user control
3. Click the button to raise the event
4. VB6 code creates a new `MyTestClass`
5. At this point it will show a message box with "test" if it works, or an automation error if it fails.


## Creating a Build to produce the error

### Prerequisites
You must have Interop Forms Toolkit installed to build the .NET user controls

### Building and running
1. Open ```DotNetControls.csproj``` and build. The post build commands copy the files and manifests to the ```build``` folder in the root of the project.
2. Open the InteropTest.vbp project Make the exe, putting it in the ```build``` folder.
3. Copy all 4 files to various machines for testing. **Don't use the dev machine without unregistering the dll first.** It works on everything but Windows 7 SP1 with updates for me.

