'''

Copyright 2017 Debasish Mandal

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

'''


import sys
try:
	sys.path.append('ExtDepLibs')
	import autoit
except:
	print('[Error] pyautoit is not installed. Which is required to run this fuzzer (Error POPUp Killer). Install pyautoit First https://pypi.python.org/pypi/PyAutoIt/0.3')
	exit()
from datetime import datetime
class PopUpKiller:
	def __init__(self):
		None
	def POPUpKillerThread(self):
		print '[+] '+ datetime.now().strftime("%Y:%m:%d::%H:%M:%S") +' POP Up killer Thread started..'
		while True:
			try:
				# MS Word
				if "Word found unreadable" in autoit.win_get_text('Microsoft Word'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "You cannot close Microsoft Word because" in autoit.win_get_text('Microsoft Word'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "caused a serious error the last time it was opened" in autoit.win_get_text('Microsoft Word'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "Word failed to start correctly last time" in autoit.win_get_text('Microsoft Word'):
					autoit.control_click("[Class:#32770]", "Button2")
				if "This file was created in a pre-release version" in autoit.win_get_text('Microsoft Word'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "The program used to create this object is" in autoit.win_get_text('Microsoft Word'):
					autoit.control_click("[Class:#32770]", "Button1") 
				if "Word experienced an error trying to open the file" in autoit.win_get_text('Microsoft Word'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "experienced an error trying to open the file" in autoit.win_get_text('Microsoft Word'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "Word was unable to read this document" in autoit.win_get_text('Microsoft Word'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "The last time you" in autoit.win_get_text('Microsoft Word'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "Safe mode could help you" in autoit.win_get_text('Microsoft Word'):
					autoit.control_click("[Class:#32770]", "Button2")
				if "You may continue opening it or perform" in autoit.win_get_text('Microsoft Word'):	
					autoit.control_click("[Class:#32770]", "Button2")  # Button2 Recover Data or Button1 Open
				#Outlook
				if "Safe mode" in autoit.win_get_text('Microsoft Outlook'):	
					autoit.control_click("[Class:#32770]", "Button2")  # Button2 Recover Data or Button1 Open
				if "Your mailbox has been" in autoit.win_get_text('Microsoft Exchange'):	
					autoit.control_click("[Class:#32770]", "Button2")  # Button2 Recover Data or Button1 Open
				
				# MS Excel 
				if "Word found unreadable" in autoit.win_get_text('Microsoft Excel'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "You cannot close Microsoft Word because" in autoit.win_get_text('Microsoft Excel'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "caused a serious error the last time it was opened" in autoit.win_get_text('Microsoft Excel'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "Word failed to start correctly last time" in autoit.win_get_text('Microsoft Excel'):
					autoit.control_click("[Class:#32770]", "Button2")
				if "This file was created in a pre-release version" in autoit.win_get_text('Microsoft Excel'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "The program used to create this object is" in autoit.win_get_text('Microsoft Excel'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "because the file format or file extension is not valid" in autoit.win_get_text('Microsoft Excel'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "The file you are trying to open" in autoit.win_get_text('Microsoft Excel'):	
					autoit.control_click("[Class:#32770]", "Button1")
				if "The file may be corrupted" in autoit.win_get_text('Microsoft Excel'):
					autoit.control_click("[Class:#32770]", "Button2")
				if "The last time you" in autoit.win_get_text('Microsoft Excel'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "We found" in autoit.win_get_text('Microsoft Excel'):
					autoit.control_click("[Class:#32770]", "Button1")
					
				#PPT
				if "The last time you" in autoit.win_get_text('Microsoft PowerPoint'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "PowerPoint found a problem with content"  in autoit.win_get_text('Microsoft PowerPoint'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "read some content" in autoit.win_get_text('Microsoft PowerPoint'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "Sorry" in autoit.win_get_text('Microsoft PowerPoint'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "PowerPoint" in autoit.win_get_text('Microsoft PowerPoint'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "is not supported" in autoit.win_get_text('SmartArt Graphics'):
					autoit.control_click("[Class:#32770]", "Button2")
				if "Safe mode" in autoit.win_get_text('Microsoft PowerPoint'):	
					autoit.control_click("[Class:#32770]", "Button2")  # Button2 Recover Data or Button1 Open
					
				# Outlook
				
				# XPS Viewer
				if "Close" in autoit.win_get_text('XPS Viewer'):
					autoit.control_click("[Class:#32770]", "Button1")
				if "XPS" in autoit.win_get_text('XPS Viewer'):
					autoit.control_click("[Class:#32770]", "Button1")
				autoit.win_close('[CLASS:bosa_sdm_msword]')
			except:
				pass
