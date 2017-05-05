'''

Copyright 2017 Debasish Mandal

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

'''
import os
from collections import namedtuple
import re
from time import sleep
def pprinttable(rows):
  # Source : http://stackoverflow.com/questions/5909873/how-can-i-pretty-print-ascii-tables-with-python
  if len(rows) > 1:
    headers = rows[0]._fields
    lens = []
    for i in range(len(rows[0])):
      lens.append(len(max([x[i] for x in rows] + [headers[i]],key=lambda x:len(str(x)))))
    formats = []
    hformats = []
    for i in range(len(rows[0])):
      if isinstance(rows[0][i], int):
        formats.append("%%%dd" % lens[i])
      else:
        formats.append("%%-%ds" % lens[i])
      hformats.append("%%-%ds" % lens[i])
    pattern = " | ".join(formats)
    hpattern = " | ".join(hformats)
    separator = "-+-".join(['-' * n for n in lens])
    print hpattern % tuple(headers)
    print separator
    _u = lambda t: t.decode('UTF-8', 'replace') if isinstance(t, str) else t
    for line in rows:
        print pattern % tuple(_u(t) for t in line)
  elif len(rows) == 1:
    row = rows[0]
    hwidth = len(max(row._fields,key=lambda x: len(x)))
    for i in range(len(row)):
      print "%*s = %s" % (hwidth,row._fields[i],row[i])
	  
class CrashSummary:
	def __init__(self):
		pass
	def getCrashDetails(self,path):
		files = os.listdir('crashes\\'+path)
		for file in files:
			if file.split('.')[-1] == 'txt':
				f = open('crashes\\'+path+'\\'+file)
				lines = f.readlines()
				func_line = lines[0]
				exp_line = lines[2]
				f.close()
				match1 = re.search(r'at .+',func_line)
				#match2 = re.search(r'at .+',exp_line)
				if match1:
					return match1.group(),exp_line
		return 'N/A'
	def Summarize(self):
		print '-----------------'
		print 'Crash Summary   |'
		print '-----------------'
		Row = namedtuple('Row',['Application','FaultAddress','UniqueCrashCount','FunctionName','Exploitability'])
		data = Row(' ',' ',' ',' ',' ')
		rows = [data]
		apps = os.listdir('Crashes')
		if len(apps) == 0:
			print 'No Crash So far :( :('
		else:
			for app in apps:
				addres = os.listdir('Crashes\\'+app)
				UniqueCrashCount = len(addres)
				for FaultAddress in addres:
					name,exploitability = self.getCrashDetails(app+'\\'+FaultAddress)
					data = Row(app,FaultAddress,UniqueCrashCount,name,exploitability)
					rows.append(data)
			pprinttable(rows)
		print '\n\nRefresh in 2 seconds.....'
if __name__ == "__main__":
	print 
	cs = CrashSummary()
	while True:
		cs.Summarize()
		sleep(2)
		os.system('cls')