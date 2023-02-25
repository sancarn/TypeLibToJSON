# TypeLibToJSON

This repository contains numerous OLE Type Libraries in JSON format. These JSON files can be used to:

* For meta-programming while using OLE, for instance if you want to automate `Excel.Application` you can use `Excel.json` to obtain runtime (or compile-time) resources and type information.
* For building IDE tools (like intellisense / type information in VSCode plugins)
* For building VBA wrapper libraries (e.g. wrap Excel object library with stdVBA, to return stdEnumerators instead of collections etc.)
* For library/data analysis (e.g. finding out how many Classes don't have a `Parent` property etc.)

You can find exported JSON files in the [`/TypeLibs/`](https://github.com/sancarn/TypeLibToJSON/tree/master/TypeLibs) folder.

## Exporter config

TBC

## Installation

### On 64-bit systems

Press start, search for and open `Component Services`.

<!-- ![]() -->

Open path `Console Root>Component Services>Computers>My Computer>COM+ Applications`. Right click and select `New>Application`
![A](/Resources/Installation_A.png)


Click `Next` and `Create an empty application`.

![B](/Resources/Installation_B.png)


Enter the name `tlbinf`. Ensure `Server application` is selected.

![C](/Resources/Installation_C.png)

Click `Next` and it is suggested you use `Interactive User` (Note: this is where you select which users can use the library. Interactive user will be any user with the correct permissions set in the next steps. If you want to install for all users simply select `Interactive User` and keep clicking `next` til the window closes.)

![D](/Resources/Installation_D.png)

You should now see the `tlbinf` application in `COM+ Applications`.

![E](/Resources/Installation_E.png)

Open `tlbinf` and right click on `Components`. Select `New>Component`.

![F](/Resources/Installation_F.png)

Click `Next` and `Install new component(s)`.

![G](/Resources/Installation_G.png)

Browse to the `TLBINF32.dll` file downloaded with this repo. Select it and click `open`

![H](/Resources/Installation_H.png)

Click `Next` and `Finish`. Now you can close `Component Services` and open the Excel tool to run it!

### On 32-bit systems

Open `cmd.exe` or `Powershell.exe`

Run the command `RegSvr32 {path}` and replace `{path}` with the full file path to the enclosed `TLBINF32.dll` file.