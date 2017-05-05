'''

Copyright 2017 Debasish Mandal

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

'''

from OfficeFileProcessor import OfficeFileProcessor
import sys
import os
class OXDumper:
	def __init__(self):
		pass
	def getFiles(self):
		files = os.listdir('BaseOfficeDocs\\OpenXMLFiles')
		if len(files) == 0:
			print '[+] No OpenXML file found in base folder :  BaseOfficeDocs\\OpenXMLFiles'
			exit()
		else:
			all_internal_files = []
			for file in files:
				op = OfficeFileProcessor('BaseOfficeDocs\\UnpackedMSOpenXMLFormatFiles','',True,'BaseOfficeDocs\\OpenXMLFiles')
				ox_dict = op.generateDict(file.split('\\')[-1])
				all_internal_files = all_internal_files + ox_dict.keys()
				#print ox_dict.keys()
			return all_internal_files
		'''
		elif len(files) > 1:
			#print len(files)
			print '[+] There are multiple files in folder : BaseOfficeDocs\\OpenXMLFiles. Please provide only one OpenXML file as base when targetting specific file(s) inside openXML document.'
			exit()
		'''
		#else
		#	op = OfficeFileProcessor('BaseOfficeDocs\\UnpackedMSOpenXMLFormatFiles','',True,'BaseOfficeDocs\\OpenXMLFiles')
		#	ox_dict = op.generateDict(files[0].split('\\')[-1])
		#	return ox_dict.keys()
if __name__ == "__main__":
	if len(sys.argv) != 2:
		print '[+] Accepts only one command line argument(comma separated file extensions). Usage : OXDumper.py xml,rels,ext1,ext2'
		od = OXDumper()
		files = od.getFiles()
		print '################ You can simply copy paste follwing python file list to config file #####################'
		list_buff = 'FILES_TO_BE_FUZZED = ['
		for file in files:
			list_buff +=  'r\'' + file  +'\',\n'
		list_buff = list_buff[:-2]+']'
		print list_buff
		print '################ File list ends #####################'
	else:
		exts = (sys.argv[1]).split(',')
		if len(exts) == 0:
			print '[+] [Error]You must provide at least one extension. c:\fuzzer>OXDumper.py xml'
			exit()
		od = OXDumper()
		files = od.getFiles()
		if len(files) == 0:
			print '[+] No files found inside Base OpenXML document with provided exntension(s) :( ',exts
			exit()
		print '################ You can simply copy paste follwing python file list to config file #####################'
		list_buff = 'FILES_TO_BE_FUZZED = ['
		for file in files:
			if file.split('.')[-1] in exts:
				list_buff +=  'r\'' + file  +'\',\n'
		list_buff = list_buff[:-2]+']'
		print list_buff
		print '################ File list ends #####################'