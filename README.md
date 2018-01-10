i) OpenXMolar v 1.0
=======================

![alt text](https://github.com/debasishm89/OpenXMolar/blob/master/screens/logo.PNG "OpenXMolar")

OpenXMolar is a [Microsoft Open XML](https://en.wikipedia.org/wiki/Office_Open_XML) file format fuzzing framework, written in Python. 


ii) Motivation Behind OpenXMolar
=========================================
MS OpenXML office files are widely used and the attack surface is huge, due to complexity of the softwares that supports OpenXML format. Office Open XML files are zipped, XML-based file format. I could not find any easy to use OpenXML auditing tools/framework available on the internet which provides software security auditors a easy to use platform using which auditors can write their own test cases and tweak internal structure of Open XML files and run fuzz test (Example : Microsoft Office).

Hence OpenXMolar was developed, using which software security auditors can focus, only on writing test cases for tweaking OpenXML internal (XML and other ) files and the framework takes care of rest of the things like unpacking, packing of OpenXML files, Error handling, etc.

iii) Dependencies
==================
OpenXMolar is written and tested on Python v2.7. OpenXMolar uses following third party libraries

1. [winappdbg](https://github.com/MarioVilas/winappdbg) / [pydbg](https://github.com/OpenRCE/pydbg)

Debugger is an immense part of any Fuzzer. Open X-Molar supports two python debugger, one is winappdbg and another is pydbg. Sometimes installing pydbg on windows environment can be painful, and pydbg code base is not well maintained hence winappdbg support added to Open X-Molar. Its recommended that user use winappdbg.

2. [pyautoit](https://pypi.python.org/pypi/PyAutoIt/0.3)

Since we feed random yet valid data into target application during fuzzing, target application reacts in many different ways. During fuzzing the target application may throw different errors through different pop-up windows. To continue the fuzzing process, the fuzzer must handle these pop-up error windows properly. OpenXMolar uses PyAutoIT to suppress different application pop-up windows. PyAutoIt is Python binding for AutoItX3.dll

3. [crash_binning.py](https://github.com/OpenRCE/sulley)

crash_binning is part of sulley framework. crash_binning.py is used only when you've selected pydbg as debugger. crash_binning.py is used to dump crash information. This is only required when you are using pydbg as debugger. 

4. [xmltodict](https://github.com/martinblech/xmltodict)

This is not core part of the Open X-Molar. The XML String Mutation module (FileFormatHandlers\xmlHandler.py) was written using xmltodict library. 



iv) Architecture:
=================

On a high level, OpenXMolar can be divided into few components.

1. OpenXMolar.py

This is the core component of this Tool and responsible for doing many important stuffs like the main fuzzing loop.

2. OfficeFileProcessor.py

This component mostly handles processing of OpenXML document such as packing, unpacking of openxml files, mapping them in memory, converting OpenXML document to python data structures etc. 


3. PopUpKiller.py - PopUp/Error Message Handlers : 

This component suppresses/kills unwanted pop-ups appeared during fuzzing. 

4. FileFormatHandlers//

An OpenXML file may contain various files like XML files, Binary files etc. FileFormatHandlers are basically a collection of mutation scripts, responsible for handling different files found inside an OpenXML document and mutate them.


5. OXDumper.py

OXDumper.py decompresses OpenXML files provided in folder "OpenXMolar\\BaseOfficeDocs\\OpenXMLFiles" and output a python list of files present in the OpenXML file. OXDumper.py accepts comma separated file extensions.
OXDumper.py is useful when you are targeting any specific set of files present in any OpenXML document.

6. crashSummary.py

crashSummary.py summarizes crashes found during fuzzing process in tabular format. The output of crashSummary.py should look like this:

![alt text](https://github.com/debasishm89/OpenXMolar/blob/master/screens/CrashSummary.png "CrashSummary")



v) Configuration File Walk through
=================================
The default configuration file '[config.py](https://github.com/debasishm89/OpenXMolar/blob/master/config.py)' is very well commented and explains all of its parameters really well. Please review the default config.py file thoroughly before running the fuzzer to avoid unwanted errors.


vi) Writing your Open XML internal File Mutation Scripts:
==========================================================

As said earlier, an OpenXML file package may contain various files like XML files, Binary files etc. FileFormatHandlers are basically a collection of mutation scripts, responsible for handling different files found inside an OpenXML document and mutate them. Generating effective test cases is the most important step in any fuzz testing process. 

The motive behind OpenXMolar was to provide security auditors an easy & flexible platform on which fuzz tester can write their own test cases very easily for OpenXML files. When it comes to effective OpenXML format fuzzing, the main part is how we mutate different files (*.xml, *.bin etc) present inside OpenXML package (zip alike). To give users an idea of how file format handlers are written, two file format handlers are provided with this fuzzer, however they are very dumb in nature and not very effective.

| FileHandler | Description |
| ------ | ------ |
| xmlHandler.py | Responsible for mutation of XML files. |
| binaryHandler.py  | Responsible for mutation of binary format files. |
| SampleHandler.py | Blank file format handler template |


Any file format handler module should be of following structure 
```python
# Import whatever you want.
class Handler():# The class name should be always 'Handler'
	def __init__(self):
		pass
	def Fuzzit(self,actual_data_stream):	
		# A function called Fuzzit must be present in Handler class
		# and it should return fuzzed data/xml string/whatever.
		# Note: Data type of actual_data_stream and data_after_mutation should always be same.

		return data_after_mutation

```

Once your file format handler module is ready you need to place the *.py file in FileFormatHandlers// folder and add the handler entry and associated file extension in config.py file like this : 

```python
FILE_FORMAT_HANDLERS = {'xml':'xmlHandler.py',
						'bin':'BinaryHandler.py',
						'rels':'xmlHandler.py',
						'vml':'xmlHandler.py'
						}
```						

vii)Adding More POPUP / Errors Windows Handler
===============================================

The default PopUpKiller.py file provided with Open X-Molar, is having few most occurred pop up / error windows handler for MS Word, MS Excel & Power Point. Using AutoIT Window Info tool (https://www.autoitscript.com/site/autoit/downloads/) you can add more POPUP / Errors Windows Handlers into 'PopUpKiller.py'. One example is given below.

![alt text](https://github.com/debasishm89/OpenXMolar/blob/master/screens/popuphandler.PNG "OpenXMolar")

So to be able to Handle the error pop up window shown in screen shot, following lines need to be added in : PopUpKiller.py

```python
if "PowerPoint found a problem with content"  in autoit.win_get_text('Microsoft PowerPoint'):
	autoit.control_click("[Class:#32770]", "Button1")

```



viii)The First Run
===================
This fuzzer is well tested on 32 Bit and 64 Bit Windows Platforms (32 Bit Office Process). All the required libraries are distributed with this fuzzer in 'ExtDepLibs/' folder. Hence if you have installed python v2.7, you are good to go. 

To verify everything is at right place, better to run Open X-Molar with Microsoft Default XPS Viewer first time(C:\\Windows\\System32\\xpsrchvw.exe). Place any *.oxps file in '\BaseOfficeDocs\OpenXMLOfficeFiles' and run OpenXMolar.py.

OpenXMolar.py accepts one command line argument which is the configuration file.

```sh
C:\Users\John\Desktop\OpenXMolar>python OpenXMolar.py config.py

[Warning] Pydbg was not found. Which is required to run this fuzzer. Install Pydbg First. Ignore if you have winappdbg installed.

   ____                    __   ____  __       _
  / __ \                   \ \ / /  \/  |     | |
 | |  | |_ __   ___ _ __    \ V /| \  / | ___ | | __ _ _ __
 | |  | | '_ \ / _ \ '_ \    > < | |\/| |/ _ \| |/ _` | '__|
 | |__| | |_) |  __/ | | |  / . \| |  | | (_) | | (_| | |
  \____/| .__/ \___|_| |_| /_/ \_\_|  |_|\___/|_|\__,_|_|
        | |
        |_|
        An MS OpenXML File Format Fuzzing Framework.
        Author : Debasish Mandal (twitter.com/debasishm89)

[+] 2017:05:05::23:11:23 Using debugger :  winappdbg
[+] 2017:05:05::23:11:23 POP Up killer Thread started..
[+] 2017:05:05::23:11:24 Loading base files in memory from :  BaseOfficeDocs\UnpackedMSOpenXMLFormatFiles
[+] 2017:05:05::23:11:24 Loading File Format Handler for extension :  xml => xmlHandler.py
[+] 2017:05:05::23:11:24 Loading File Format Handler for extension :  rels => xmlHandler.py
[+] 2017:05:05::23:11:24 Loading File Format Handler Done !!
[+] 2017:05:05::23:11:24 Starting Fuzzing
[+] 2017:05:05::23:11:25 Temp cleaner started...
[+] 2017:05:05::23:11:25 Cleaning Temp Directory...
...
...
```


ix) Open X-Molar in Action
==========================

Here is a very short video on running fuzztest on MS Office Word: 

[![IMAGE ALT TEXT HERE](https://github.com/debasishm89/OpenXMolar/blob/master/screens/video_thumb.PNG)](https://www.youtube.com/watch?v=b7n1tuFDl5A)

x) Fuzzing Non-OpenXML Applications :
====================================

Due to the flexible structure of the fuzzer, this Fuzzer can also be used to fuzz other windows application. You just need do following :


* In config.py add the target application binary (exe) and extension in APP_LIST of config.py
* In config.py change OpenXMLFormat to False
* Write your own File format mutation handler and place it in FileFormatHandlers/ folder
* Add the newly added FileFormatHandler in FILE_FORMAT_HANDLERS of config.py 
* Provide some base files in folder OtherFileFormats/
* Add custom error / popup windows handler in PopUpKiller.py using Au3Info tool if required

And you're good to go.

xi) Few More Points about OpenXMolar:
======================================
1. Fuzzing Efficiency:
To maximize fuzzing efficiency OpenXMolar doesn't read the provided base files again and from disk. While starting up, it loads all base files in memory and convert them into easy to manage python data structures and mutate them straight from memory.

2. Auto identification of internal files of OpenXML package :
An Open XML file package may contain various files like XML files, Binary files etc. OpenXMolar has capability to identify internal file types and based that chooses mutation script and mutate them. Please refer to the default config.py file (Param : AUTO_IDENTIFY_INTERNAL_FILE_FORAMT) for details.


xii) TODO
=======
1. Improve Fuzzing Speed
2. New Feature / Bugs -> https://github.com/debasishm89/OpenXMolar/issues

xiii) Licence
===========

This software is licenced under New BSD License although the following libraries are included with Open X-Molar and are licensed separately.

| Module | Source |
| ------ | ------ |
| winappdbg | https://github.com/MarioVilas/winappdbg |
| pydbg | https://github.com/OpenRCE/pydbg |
| pyautoit | https://pypi.python.org/pypi/PyAutoIt/0.3 |
| crash_binning | https://github.com/OpenRCE/sulley |
| xmltodict | https://github.com/martinblech/xmltodict |
| Au3Info.exe | https://www.autoitscript.com/autoit3/docs/intro/au3spy.htm |


xiv )Author
=============

Debasish Mandal ( https://twitter.com/debasishm89 )

xv ) CVE(s)
============
https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/CVE-2017-8630
https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/CVE-2018-0792
