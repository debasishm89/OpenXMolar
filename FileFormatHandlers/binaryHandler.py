from random import randrange
from math import ceil
from random import randint
from random import choice

class Handler():
	def __init__(self):
		pass
	def int2binary(self,n, cnt=24):
		return "".join([str((n >> y) & 1) for y in range(cnt-1, -1, -1)])
	def Fuzzit(self,data):
		'''
		Fuzz Binary data.
		'''
		fuzzratio = float(0.09)
		b = list(data)
		data_to_be_fuzzed = randrange(ceil((float(len(data))) * fuzzratio))+1		#Number of bytes to fuzz
		case = randint(0,1)
		if case == 0:
			# Replace random offsets with random chars
			for j in range(data_to_be_fuzzed):		#Iterate all bytes
				randbyte = randrange(256)			#Random character
				ran = randrange(len(data))		#Random offset
				#ran = randint(10,len(data))			#Random offset Leave first 10 bytes
				b[ran] = '%c'%(randbyte)			#Replace
			mutated =''.join(b)						#Append
		if case == 1:
			# Bit flip randomly chosen bytes
			for j in range(data_to_be_fuzzed):			#Iterate
				ran = randrange(len(data))				#Random offset
				#ran = randint(10,len(data))			#Random offset Leave first 10 bytes
				bits = self.int2binary(ord(b[ran]),8)	#Int to Binary
				flipped = bits[::-1]					#Bit-flip
				b[ran] = chr(int(flipped,2))			#Replace
			mutated =''.join(b)							#Append
		return mutated	
