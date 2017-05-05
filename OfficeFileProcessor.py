'''

Copyright 2017 Debasish Mandal

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

'''

import os
import zipfile
from datetime import datetime

class OfficeFileProcessor:
	'''
	Code responsible for OpenXML file handling.
	'''
	def __init__(self,UNPACKED_OFFICE_PATH,TEMP_PATH,oxml,PACKED_OFFICE_PATH):
		self.UNPACKED_OFFICE_PATH = UNPACKED_OFFICE_PATH
		self.PACKED_OFFICE_PATH = PACKED_OFFICE_PATH
		self.TEMP_PATH = TEMP_PATH
		self.ALL_DOCS_IN_MEMORY = {}
		self.oxml = oxml
		if self.oxml:
			# its an open xml file, unpack all of them.
			files = os.listdir(self.PACKED_OFFICE_PATH)
			for file in files:
				self.ExtractOpenXML(self.PACKED_OFFICE_PATH+'\\'+file,self.UNPACKED_OFFICE_PATH+'\\'+file.replace(' ','_'))
	def Pack2OfficeDoc2(self,op_name,dict):
		'''
		
		Compress in memory dict to OpenXML file in disk.
		
		'''
		if self.oxml:
			# Its an open xml format file
			try:				# Sometimes when office responds too slowly, fuzzer may crash,hence adding this, so fuzzing continues
				zipf = zipfile.ZipFile(self.TEMP_PATH+'\\'+op_name, 'w')
				for file in dict:
					#print file
					l = len(self.UNPACKED_OFFICE_PATH + '\\' + op_name)
					zipf.writestr(file[l+1:], dict[file], zipfile.ZIP_DEFLATED ) #  encode to get rid of Unicode error.
				zipf.close()
			except:
				pass
		else:
			# Its a binary file
			#print 'packing binary files'
			#print self.TEMP_PATH + '\\' + op_name
			try:
				f = open(self.TEMP_PATH + '\\' + op_name,'wb')
				f.write(dict)		# In this case its a Binary stream
				f.close()
			except:
				pass
	def ExtractOpenXML(self,path_to_zip_file,directory_to_extract_to):
		'''
		Extract content of an OpenXML file.
		
		'''
		try:
			zip_ref = zipfile.ZipFile(path_to_zip_file, 'r')
			zip_ref.extractall(directory_to_extract_to)
			zip_ref.close()
		except:
			print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'Problem while extracting ',path_to_zip_file
	def ReadContent(self,path):
		f = open(path,'rb')
		d = f.read()
		f.close()
		return d
	def generateDict(self,dir):
		'''
		
		Generate a dict out of an OpenXML file, So that we can manage easily.
		
		'''
		office_dict = {}
		path = self.UNPACKED_OFFICE_PATH + '\\' + dir
		for root, dirs, files in os.walk(path):
			for file in files:
				office_dict[os.path.join(root, file)] = self.ReadContent(os.path.join(root, file))
		return office_dict
	def LoadFilesInMemory(self):
		'''
		
		Load all Base files in memory.
		
		'''
		if self.oxml:
			for curr in os.listdir(self.UNPACKED_OFFICE_PATH):
				#print '[+] Loading file in memory',curr
				office_dict = self.generateDict(curr)
				self.ALL_DOCS_IN_MEMORY[curr] = office_dict
			return  self.ALL_DOCS_IN_MEMORY
		else:
			for curr in os.listdir(self.UNPACKED_OFFICE_PATH):
				office_bin_cont = self.ReadContent(self.UNPACKED_OFFICE_PATH+'\\'+curr)
				self.ALL_DOCS_IN_MEMORY[curr] = office_bin_cont
			return  self.ALL_DOCS_IN_MEMORY
