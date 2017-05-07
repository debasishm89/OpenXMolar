# -*- coding: utf-8 -*- 

'''

Copyright 2017 Debasish Mandal

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

'''


# Configuration file for OpenXMolar.


# Folder where base office OpenXML files will be kept.(Example *.docx , *.xlsx etc.)
packed_open_xml_office_files = 'BaseOfficeDocs\\OpenXMLFiles'

# Folder where base office binary format files are kept. For example *.doc
binary_office_files = 'BaseOfficeDocs\\OtherFileFormats'

# Temporary folder where base office OpenXML files will be unpacked before they are loaded in memory. 
open_xml_office_files = 'BaseOfficeDocs\\UnpackedMSOpenXMLFormatFiles'

# Temporary folder where Test cases will be kept, before they are opened.
fuzztempfolder  = 'FuzzTemp'


# File format to be fuzzed.
# If OpenXMLFormat is False, Base files will be taken from folder 'binary_office_files'
# If OpenXMLFormat is True, Base files will be taken from folder 'open_xml_office_files'
OpenXMLFormat = True

####################################################################################################################
# This dictionary is holding the mapping between office application and extension. 
# While fuzzing, an application (exe) will be chosen from this dictionary based on extension of file(s) provided in folder 'open_xml_office_files' or 'binary_office_files'
# Add your own extension and exe. :). Note: Make sure extension does not repeat. 
APP_LIST = {
			'xps':r'C:\Windows\system32\xpsrchvw.exe',
			'oxps':r'C:\Windows\system32\xpsrchvw.exe'
			}
			
# A Python list of command line arguments to be used while running target application. To be left blank when none required. Example : COMMAND_LINE_ARGUMENT = ['arg1','arg2','arg3','arg4']
COMMAND_LINE_ARGUMENT = []
####################################################################################################################


########################################################################################
# An Open XML file package may contain various files like XML files, Binary files etc.  
# When AUTO_IDENTIFY_INTERNAL_FILE_FORAMT is set to 'True', OpenXMolar will try to indentify types of files present inside provided base openxml packages and based on that decide, which mutation script to to use during fuzzing.			
AUTO_IDENTIFY_INTERNAL_FILE_FORAMT = True


# Mutation script list. Once AUTO_IDENTIFY_INTERNAL_FILE_FORAMT is set to True, available mutation scripts and associated extensions has to be defined under ALL_MUTATION_SCRIPTS in dict format.

ALL_MUTATION_SCRIPTS = {'xml':'SampleHandler.py','bin':'binaryHandler.py'}
#########################################################################################


			
# Note: This is only required when AUTO_IDENTIFY_INTERNAL_FILE_FORAMT is set to False.
# Following dict. holds a mapping between different file extension and a file which can parse the format and mutate its content (handler).
# For example when the fuzzer will find a *.xml file inside an OpenXML document it will use '\FileFormatHandlers\SampleHandler.py' file to mutate the xml file.
# While writing your own custom file format handler , please refer sample : \FileFormatHandlers\SampleHandler.py and README.md doc.
# File format handler(s) should always be kept inside '\FileFormatHandlers\' folder.

FILE_FORMAT_HANDLERS = {'xml':'SampleHandler.py','rels':'SampleHandler.py'}


						
# Delay Between test case iteration;
FUZZ_LOOP_DELAY = 0.5 	# In Seconds


# Run the application for n seconds, and monitor it.
# Tip: MSOffice usually considered as heavy application. If you are not running this fuzzer on a very fast system, increase APP_RUN_TIME accordingly.
APP_RUN_TIME = 1.5 	# In seconds


# Debugger to use for process monitoring.
# Right now MSOXFuzz supports two debuggers, 'winappdbg' and 'pydbg'. Installing pydbg could be painful sometimes, so use 'winappdbg' proudly :) It works pretty good.
DEBUGGER = 'winappdbg'   

# Number of files to be mutated inside any base OpenXML document. For better results keep it less, For example 2,3 or max 5.
NUMBER_OF_FILES_TO_MUTATE = 2


# In case if you want to target/fuzz one or more specific file(s) inside an OpenXML document, you need to provide the file name(s) in this list 'FILES_TO_BE_FUZZED' depending on base files present at packed_open_xml_office_files
# For example in following OpenXML structure, if you want to target 'MXDC_Empty_PT.xml' & 'FixedDocument.fdoc' file.
# The list would look something like this : FILES_TO_BE_FUZZED = ['MXDC_Empty_PT.xml','FixedDocument.fdoc']
# In case if you want to fuzz all files, keep it empty FILE_TO_BE_FUZZED = []

# Note: This is only required when OpenXMLFormat = True , config. file.
# You can use OXDumper.py to generate this list.
# Warning : When FILES_TO_BE_FUZZED is not empty , you must provide only one OpenXML file as base file. When provided multiple base file this list should be empty.
# Please remember to change this list when you change the Basefile in OpenXML base file folder. packed_open_xml_office_files to avoid nasty errors.

################ Generated using OXDumper.py, You can simply copy paste follwing python file list to config file #####################
FILES_TO_BE_FUZZED = []
################ File list ends #####################
'''
Sample OpenXML directory structure

│   FixedDocumentSequence.fdseq
│   [Content_Types].xml
├───Documents
│   └───1
│       │   FixedDocument.fdoc
│       ├───Metadata
│       │       Page1_Thumbnail.JPG
│       ├───Pages
│       │   │   1.fpage
│       │   └───_rels
│       │           1.fpage.rels
│       ├───Resources
│       │   └───Images
│       │           1.JPG
│       └───_rels
│               FixedDocument.fdoc.rels
├───Metadata
│       Job_PT.xml
│       MXDC_Empty_PT.xml
├───Resources
│       _D1.dict
└───_rels
        .rels
        FixedDocumentSequence.fdseq.rels

'''
