# pymsproject

This utility will parse a Microsoft Project XML file and attempt to derive
some insights into the contents.

## Generating the Project XML File
The XML file is presently required instead of directly using the Microsoft
Project file format (.MPP). In order to properly generate this file, perform
the following steps:

* Open the project file in Microsoft Project.
* Click _File_ then _Save As_.
* Using the _Save as type:_ drop down, select _XML_.
* Save the resulting XML file to known location, i.e. Desktop.

## Generating a Milestone PNG File
Once you have an updated XML file. Open this application and press the _Browse_
button to select the XML file. As soon as you have selected the file, this
application will attempt to generate the new report data.

If you are happy with the image file as presented on the screen, press the
_Save_ button to generate a PNG file!

## Working from Source
The following information is only required if you are working directly from the
source code. If you are using the executable, then the following steps are not
very interesting!

### Pre-requisites
* Ensure that you have installed Python 3+ onto your system.
* Open a command prompt and navigate to the folder containing this source.
* Execute ```pip install -r reqs.txt``` to install the required Python libraries.
* If you are planning on building a .exe bundle, also execute:
```pip install pyinstaller```

### Building the source
* Open a command prompt and navigate to the folder containing this source.
* Execute the following command:
```
pyinstaller --onefile --noconsole app.py
```
* An executable binary will be placed into the _dist_ folder of the project.
