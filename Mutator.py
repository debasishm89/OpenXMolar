'''

Copyright 2017 Debasish Mandal

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

'''

import os
from random import randint
from random import choice
from random import randrange
from collections import OrderedDict
import imp
from datetime import datetime
from math import ceil
from copy import copy



class Mutator:
	'''
	
	This class is responsible for data mutation (depending on xml / binary format it will choose appropriate file format mutation hanlder) . Its uses provided file handlers to mutate data.
	
	'''
	def __init__(self,OpenXML,handlers,files_to_be_fuzzed,no_of_files_to_be_fuzzed):
		self.oxml = OpenXML
		self.HANDLERS = handlers
		self.handler_obj_dict = {}
		self.LoadFormatHandlers()
		self.FILES_TO_BE_FUZZED = files_to_be_fuzzed
		self.NUMBER_OF_FILES_TO_MUTATE = no_of_files_to_be_fuzzed
		
	def LoadFormatHandlers(self):
		for handler in self.HANDLERS:
			try:
				print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'Loading File Format Handler for extension : ',handler,'=>',self.HANDLERS[handler]
				foo = imp.load_source('Handler', 'FileFormatHandlers//'+self.HANDLERS[handler])
				a = foo.Handler()
				self.handler_obj_dict[handler] = a
			except:
				print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'There is an error in this Fileformat handler or it was not written correctly.','FileFormatHandlers//'+self.HANDLERS[handler], 'Please check FileFormatHandlers\\SampleHandler.py'
		print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'Loading File Format Handler Done !!'
	def Mutate(self,office_doc_dict,file_name):
		'''
		
		This function accepts an office doc, converted into a python dictionary.
		It decides which files of the office document to be fuzzed, Based on the file (xml or binary file)		
		
		'''
		if self.oxml:
			# The content sent here is OpenXML format. Hence we are
			fuzzed_office_file_dict = {}
			fuzzed_office_file_dict = copy(office_doc_dict)
			#NUMBER_OF_FILES_TO_MUTATE = 5
			if len(self.FILES_TO_BE_FUZZED) == 0:
				# Fuzz some files randomly choosen from entire open xml file.
				for file_count in range(0,self.NUMBER_OF_FILES_TO_MUTATE):
					target_file  = choice(office_doc_dict.keys())	# The file
					ext = target_file.split('.')[-1]				# Get the extension of the file.
					if ext in self.handler_obj_dict:				# Check if format handler for the current format is present 
						HandlerClass = self.handler_obj_dict[ext]	# Get the File handler object from handler dict.
						fuzzed_office_file_dict[target_file] = HandlerClass.Fuzzit(office_doc_dict[target_file])
					else:
						fuzzed_office_file_dict[target_file] = office_doc_dict[target_file]		# Do not do anything.
			else:
				# Fuzz some files randomly chosen from the config. 
				#NUMBER_OF_FILES_TO_MUTATE = 2
				for file_count in range(0,self.NUMBER_OF_FILES_TO_MUTATE):
					target_file  = choice(self.FILES_TO_BE_FUZZED)
					#print target_file
					ext = target_file.split('.')[-1]								# Find the extension of the file
					if ext in self.handler_obj_dict:								# Check if format handler for the current format is present 
						HandlerClass = self.handler_obj_dict[ext]					# Get the handler object
						try:
							fuzzed_office_file_dict[target_file] = HandlerClass.Fuzzit(office_doc_dict[target_file])
						except:
							print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'[Error] File not found in provided base openXML file',target_file,'Please check config file parameter FILES_TO_BE_FUZZED. You can Re-Generate FILES_TO_BE_FUZZED list using OXDumper.py'
							exit()
					else:
						fuzzed_office_file_dict[target_file] = office_doc_dict[target_file]		# Do not do anything.					
			return fuzzed_office_file_dict
		else:
			# Return a binary stream
			ext = file_name.split('.')[-1]
			HandlerClass = self.handler_obj_dict[ext]
			return self.handler_obj_dict[ext].Fuzzit(office_doc_dict)