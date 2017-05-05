# Import whatever you want.



class Handler():		# The class name should be always Handler
	def __init__(self):
		pass			# Not required.

	def Fuzzit(self,actual_data):	
		# A function called Fuzzit must be present in Handler class and it should return fuzzed data/xml string/whatever.
		# Note: Date type of actual_data and data_after_mutation should always be same.
		'''
		Fuzz the content of actual_data and return the fuzzed binary data Binary data.
		'''
		data_after_mutation = actual_data
		return data_after_mutation
