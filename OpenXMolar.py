'''

Copyright 2017 Debasish Mandal

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

'''

from OfficeFileProcessor import OfficeFileProcessor
from Mutator import Mutator
from PopUpKiller import PopUpKiller

import shutil
from random import choice
from time import sleep
from datetime import datetime
import os
import os.path
import importlib
import sys
import thread
import subprocess
try:
	from pydbg import *
	from pydbg.defines import *
except:
	print('[Warning] Pydbg was not found. Which is required to run this fuzzer. Install Pydbg First. Ignore if you have winappdbg installed.')
try:
	from winappdbg import Crash,win32,Debug
except:
	print('[Error] winappdbg could not be imported. Which is required to run this fuzzer. Install winappdbg First')
	exit()

import utils

class Fuzzer:
	def __init__(self,fuzzer_config_obj):
		self.conf = self.ConfigParse(fuzzer_config_obj)
		self.CurrentTestCaseName = ''
		self.OpenXMLFormat= fuzzer_config_obj.OpenXMLFormat
		self.debugger = fuzzer_config_obj.DEBUGGER
		self.basic_cmd_arg = fuzzer_config_obj.COMMAND_LINE_ARGUMENT
		self.extra_cmd_arg = fuzzer_config_obj.COMMAND_LINE_ARGUMENT
		self.StartUPBrooming()
	def StartUPBrooming(self):
		# Do some clean up while starting up. It just removed the unpacked office file folder.
		
		contents = [os.path.join(self.conf.open_xml_office_files, i) for i in os.listdir(self.conf.open_xml_office_files)]
		[os.remove(i) if os.path.isfile(i) or os.path.islink(i) else shutil.rmtree(i) for i in contents]
		sleep(2)
		return True
	def DeleteOfficeHistorty(self):
		# Not required now.
		
		#print '[+] Deleting Safe Mode Prompt Office History'
		s = 'REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Resiliency\StartupItems" /f'
		#s = 'REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\File MRU" /v "Item 1" /f'
		os.popen(s)
	def CheckENV(self):
		# Check Operating System Type (not in use)
		if 'PROGRAMFILES(X86)' in os.environ:
			print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'Currently 64bit OS is not supported(Pydbg only supports x86).Please run it on 32 bit Windows'
			exit()
	def TempDirCleanThread(self):
		'''
		This thread does some cleanup every 15 mins.
		'''
		sleep(2)
		print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'Temp cleaner started...'
		while True:
			print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'Cleaning Temp Directory...'
			subprocess.Popen('rmdir /q /s %temp%', shell=True, stdout = subprocess.PIPE, stderr= subprocess.PIPE)
			sleep(60*15)
	def ForceKillOffice(self):
		'''
		In case debugger is unable to kill the half dead office process, we will try to kill it forcefully.
		
		'''
		print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'Forcefully Killing Office Application'
		progname = self.conf.APP_LIST[self.CurrentTestCaseName.split('.')[-1]].split('\\')[-1]
		os.popen('taskkill /PID '+ progname +' /f')
	def AccessViolationHandlerPYDBG(self,dbg):
		'''
		Handle access violation / guard page access violation using pydbg
		'''
		crash_bin = utils.crash_binning.crash_binning()		
		crash_bin.record_crash(dbg)
		violation_addr = hex(dbg.dbg.u.Exception.ExceptionRecord.ExceptionInformation[1])
		thetime = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
		crashfilename = 'crash_'+self.CurrentTestCaseName+'_'+'_'+ thetime +'.'+self.CurrentTestCaseName.split('.')[-1]
		synfilename = 'crashes\\'+ violation_addr +'\\crash_'+self.CurrentTestCaseName+'_'+'_'+ thetime +'.txt'
		if not os.path.exists('crashes\\'+violation_addr):
			os.makedirs('crashes\\'+violation_addr)
		shutil.copyfile(self.conf.fuzztempfolder + '\\' + self.CurrentTestCaseName,'crashes\\'+violation_addr+'\\'+crashfilename)
		print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'BOOM!! APP Crashed :',dbg.disasm(dbg.dbg.u.Exception.ExceptionRecord.ExceptionAddress)[0],'Crash file Copied to ',(violation_addr+'\\'+crashfilename)
		syn = open(synfilename,'w')
		syn.write(crash_bin.last_crash_synopsis())
		syn.close()
		if dbg.debugger_active:
			dbg.terminate_process()
		#self.DeleteOfficeHistorty()
		return DBG_CONTINUE
	def AccessViolationHandlerWINAPPDBG(self,event):
		
		# Handle access violation while using winappdbg

		code = event.get_event_code()
		if event.get_event_code() == win32.EXCEPTION_DEBUG_EVENT and event.is_last_chance():
			crash = Crash(event)
			crash.fetch_extra_data(event)
			details = crash.fullReport(bShowNotes=True)
			violation_addr = hex(crash.registers['Eip'])
			thetime = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
			exe_name =  event.get_process().get_filename().split('\\')[-1]
			crashfilename = 'crash_'+self.CurrentTestCaseName+'_'+'_'+ thetime +'.'+self.CurrentTestCaseName.split('.')[-1]
			synfilename = 'crashes\\'+exe_name+'\\'+ violation_addr +'\\crash_'+self.CurrentTestCaseName+'_'+'_'+ thetime +'.txt'
			if not os.path.exists('crashes\\'+exe_name):
				os.makedirs('crashes\\'+exe_name)
			if not os.path.exists('crashes\\'+exe_name+'\\'+violation_addr):
				os.makedirs('crashes\\'+exe_name+'\\'+violation_addr)
			shutil.copyfile(self.conf.fuzztempfolder + '\\' + self.CurrentTestCaseName,'crashes\\'+exe_name+'\\'+violation_addr+'\\'+crashfilename)
			print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'BOOM!! APP Crashed :','Crash file Copied to ',(exe_name+'\\'+violation_addr+'\\'+crashfilename)
			syn = open(synfilename,'w')
			syn.write(details)
			syn.close()
			#print '[+] '+ datetime.now().strftime("%Y:%m:%d::%H:%M:%S")+' Killing half dead process'
			try:
				event.get_process().kill()
			except:
				self.ForceKillOffice()
		#self.DeleteOfficeHistorty()		
		
	def StillRunningPYDBG(self,dbg):
		# This function (run as thread) kill the process after user defined interval.(usen pydbg in use)
		sleep(self.conf.APP_RUN_TIME)
		if dbg.debugger_active:
			try:
				dbg.terminate_process()
			except:
				self.ForceKillOffice()

	def StillRunningWINAPPDBG(self,proc):
		# This function (run as thread) kill the process after user defined interval.(usen winappdbg in use)
		sleep(self.conf.APP_RUN_TIME)
		try:
			proc.kill()
		except:
			self.ForceKillOffice()
	def ConfigParse(self,config_obj):
		# Parse the configuration file and check integrity of the config. file. 
		
		#1. Check if all exe file exist or not
		for ext in config_obj.APP_LIST:
			if not os.path.exists(config_obj.APP_LIST[ext]):
				print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'[Error] Application does not exist, revise APP_LIST in config file. You can remove this app. if you dont need this : ',config_obj.APP_LIST[ext]
				exit()
		
		if config_obj.OpenXMLFormat:
			#2. check if any base file provided or not
			exts = []
			files = os.listdir(config_obj.packed_open_xml_office_files)
			if len(files) == 0:
				print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'[Error] Base file directory is empty. Please provide one or more base files in directory ',config_obj.packed_open_xml_office_files
				exit()
			else:
				#5. multiple extension check, if the target folder has multiple extension.
				for file in files:
					exts.append(file.split('.')[-1])
				#print set(exts)
				if len(set(exts)) > 1:
					print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'You have provided open xml base files with mixed extensions. Please provide files of same type in \BaseOfficeDocs\OpenXMLOfficeFiles'
					exit()
				else:
					# Check if all handlers are provided in config file
					for ext in set(exts):
						if ext not in config_obj.APP_LIST.keys():
							print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'No file handler found for extension : ',ext,'. Please add file handler in config file APP_LIST'
							exit()
			#6. If multiple files provided, make sure the files_to_be_fuzzed is empty.
			if len(files) > 1 and len(config_obj.FILES_TO_BE_FUZZED) > 0:
				print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'You have provided multiple openxml base files and FILES_TO_BE_FUZZED in config is not empty. Please set FILES_TO_BE_FUZZED = [] when you have multiple open xml base files'
				exit()
		else:
			# Not an Openxml file.
			exts = []
			files = os.listdir(config_obj.binary_office_files)
			if len(files) == 0:
				print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'[Error] Base file directory is empty. Please provide one or more base files in directory',config_obj.binary_office_files
				exit()
			else:
				#5. multiple extension check, if the target folder has multiple extension.
				for file in files:
					exts.append(file.split('.')[-1])
				#print set(exts)
				if len(set(exts)) > 1:
					print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'You have provided base files with mixed extensions. Please provide files of same type \BaseOfficeDocs\BinaryOfficeFiles'
					exit()
				else:
					# Check if all handlers are provided in config file
					for ext in set(exts):
						if ext not in config_obj.APP_LIST.keys():
							print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'Base file directory is : ',config_obj.binary_office_files,', No file handler found for extension : ',ext,'. Please add file handler in config file APP_LIST'
							exit()
						# Since its not an OpenXML file, we need to check if FileFormat Mutation handler is present or not.
						if ext not in config_obj.FILE_FORMAT_HANDLERS.keys():
							print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'No file Mutation handler found for extension : ',ext,'. Please add mutation handler in config file FILE_FORMAT_HANDLERS'
							exit()
					
		#3. check if all file format handlers are present or not
		for handler in config_obj.FILE_FORMAT_HANDLERS:
			if not os.path.exists('FileFormatHandlers\\'+config_obj.FILE_FORMAT_HANDLERS[handler]):
				print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'[Error] File format handler not present, Please check FILE_FORMAT_HANDLERS in config file : ',config_obj.FILE_FORMAT_HANDLERS[handler]
				exit()
		
		
		return config_obj
	def StartFuzzer(self):
		# Main fuzzing loop
		print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'Using debugger : ',self.conf.DEBUGGER
		popup = PopUpKiller()
		thread.start_new_thread(popup.POPUpKillerThread, ())
		thread.start_new_thread(self.TempDirCleanThread, ())
		sleep(1)
		if self.OpenXMLFormat:
			print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'Loading base files in memory from : ',self.conf.open_xml_office_files
			off = OfficeFileProcessor(UNPACKED_OFFICE_PATH = self.conf.open_xml_office_files,TEMP_PATH = self.conf.fuzztempfolder,oxml=True,PACKED_OFFICE_PATH=self.conf.packed_open_xml_office_files)
			ALL_DOCS_IN_MEMORY = off.LoadFilesInMemory()
			if (len(ALL_DOCS_IN_MEMORY)) == 0:
				print '[+]', datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'No Open XML files provided as base file @',self.conf.open_xml_office_files
				exit()
			mut = Mutator(OpenXML=True,handlers = self.conf.FILE_FORMAT_HANDLERS,files_to_be_fuzzed = self.conf.FILES_TO_BE_FUZZED,no_of_files_to_be_fuzzed = self.conf.NUMBER_OF_FILES_TO_MUTATE, auto_id_file_type = self.conf.AUTO_IDENTIFY_INTERNAL_FILE_FORAMT,all_handlers=self.conf.ALL_MUTATION_SCRIPTS, all_inmem_docs = ALL_DOCS_IN_MEMORY)				# Passing file format handler and type
			print '[+]', datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'Starting Fuzzing'
			while 1:
				self.CurrentTestCaseName = choice(ALL_DOCS_IN_MEMORY.keys())					# randomly select one office file loaded in memory.				
				fuzzed_office_dict = mut.Mutate(ALL_DOCS_IN_MEMORY[self.CurrentTestCaseName],self.CurrentTestCaseName)	# fuzzed_office_dict only have xml strings in it.
				off.Pack2OfficeDoc2(self.CurrentTestCaseName,fuzzed_office_dict)				# Packed the fuzzed file back to wordfile.
				# TODO : One check can be implemented whether the files provided has proper handler application			
				PROG_NAME = self.conf.APP_LIST[self.CurrentTestCaseName.split('.')[-1]]			# Find the program name
				arg = os.getcwd() + '\\' + self.conf.fuzztempfolder + '\\' + self.CurrentTestCaseName				# 
				if self.debugger == 'pydbg':
					dbg = pydbg()
					dbg.set_callback(EXCEPTION_ACCESS_VIOLATION, self.AccessViolationHandlerPYDBG)			# AV handler
					dbg.set_callback(EXCEPTION_GUARD_PAGE, self.AccessViolationHandlerPYDBG)				# Guard Page AV handler
					thread.start_new_thread(self.StillRunningPYDBG, (dbg, ))								# Killer Thread Started
					extra_arg = " ".join(self.extra_cmd_arg)
					dbg.load(PROG_NAME,extra_arg + ' ' +arg, show_window=True)
					dbg.run()
				if self.debugger == 'winappdbg':
					self.extra_cmd_arg = list(self.basic_cmd_arg) 	# Reset it to basic
					self.extra_cmd_arg.insert(0,PROG_NAME)
					self.extra_cmd_arg.insert(len(self.extra_cmd_arg),arg)
					debug = Debug( self.AccessViolationHandlerWINAPPDBG, bKillOnExit = True )
					proc = debug.execv( self.extra_cmd_arg)
					thread.start_new_thread(self.StillRunningWINAPPDBG, (proc, ))
					debug.loop()
				sleep(self.conf.FUZZ_LOOP_DELAY)
		else:
			print '[+] '+ datetime.now().strftime("%Y:%m:%d::%H:%M:%S") +' Loading Binary files files in memory from : ',self.conf.binary_office_files
			off = OfficeFileProcessor(UNPACKED_OFFICE_PATH = self.conf.binary_office_files,TEMP_PATH = self.conf.fuzztempfolder,oxml = False,PACKED_OFFICE_PATH=self.conf.packed_open_xml_office_files)
			ALL_DOCS_IN_MEMORY = off.LoadFilesInMemory()
			if (len(ALL_DOCS_IN_MEMORY)) == 0:
				print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),' No Binary files provided as base file @',self.conf.binary_office_files
				exit()
			mut = Mutator(OpenXML=False,handlers = self.conf.FILE_FORMAT_HANDLERS,files_to_be_fuzzed = self.conf.FILES_TO_BE_FUZZED,no_of_files_to_be_fuzzed=self.conf.NUMBER_OF_FILES_TO_MUTATE,auto_id_file_type = self.conf.AUTO_IDENTIFY_INTERNAL_FILE_FORAMT,all_handlers=self.conf.ALL_MUTATION_SCRIPTS,all_inmem_docs = ALL_DOCS_IN_MEMORY)
			print '[+]', datetime.now().strftime("%Y:%m:%d::%H:%M:%S") ,'Starting Fuzzing'
			while 1:
				#print 'loop start '
				self.CurrentTestCaseName = choice(ALL_DOCS_IN_MEMORY.keys())					# randomly select one office file loaded in memory.				
				fuzzed_office_bin = mut.Mutate(ALL_DOCS_IN_MEMORY[self.CurrentTestCaseName],self.CurrentTestCaseName)	# fuzzed_office_dict only have xml strings in it.
				off.Pack2OfficeDoc2(self.CurrentTestCaseName,fuzzed_office_bin)
				PROG_NAME = self.conf.APP_LIST[self.CurrentTestCaseName.split('.')[-1]]			# FInd the program name
				arg = os.getcwd() + '\\' +self.conf.fuzztempfolder + '\\' + self.CurrentTestCaseName				# 
				if self.debugger == 'pydbg':
					dbg = pydbg()
					dbg.set_callback(EXCEPTION_ACCESS_VIOLATION, self.AccessViolationHandlerPYDBG)			# AV handler
					dbg.set_callback(EXCEPTION_GUARD_PAGE, self.AccessViolationHandlerPYDBG)				# Guard Page AV handler
					thread.start_new_thread(self.StillRunningPYDBG, (dbg, ))								# Killer Thread Started
					extra_arg = " ".join(self.extra_cmd_arg)
					dbg.load(PROG_NAME,extra_arg + ' ' +arg, show_window=True)
					dbg.run()
				if self.debugger == 'winappdbg':
					self.extra_cmd_arg = list(self.basic_cmd_arg) 	# Reset it to basic
					self.extra_cmd_arg.insert(0,PROG_NAME)			# Add progname
					self.extra_cmd_arg.insert(len(self.extra_cmd_arg),arg)
					debug = Debug( self.AccessViolationHandlerWINAPPDBG, bKillOnExit = True )
					proc = debug.execv( self.extra_cmd_arg)
					thread.start_new_thread(self.StillRunningWINAPPDBG, (proc, ))
					debug.loop()
				sleep(self.conf.FUZZ_LOOP_DELAY)

if __name__ == "__main__":
	banner = '''
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
		'''
	print banner
	if len(sys.argv) !=  2:
		print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'[Error] Config file missing..'
		print '[+]',datetime.now().strftime("%Y:%m:%d::%H:%M:%S"),'Usage : c:\>python MSOXFuzz.py config.py'
		exit()
	else:
		mod = sys.argv[1].split('.')[0]
		conf = importlib.import_module(mod, package=None)
		fz = Fuzzer(conf)
		fz.StartFuzzer()
