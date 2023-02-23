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

TBC
