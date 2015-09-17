#Python file for configuration file reading

class CConfigParser:

	def __init__(self, filename):
		self.configFile = filename   #"xlscan.conf"
		self.excel_filename = ""
		self.output_filename = ""
		self.output_all = False
		self.worksheetList = []
		self.specialCellsDict = {}
		self.ignoreList = []

	#-------------------------------------------
	# Operation codes
	# 0 - Excel workbook name
	# 1 - Name of output file
	# 2 - Output for each worksheet? True/False
	# 3 - Load worksheet names and max cells
	# 4 - Load special names and time values
	# 5 - Load the ignore list
	#-------------------------------------------
	def setValues(self, line, operation):
		bufstr = line.strip()
		
		if operation == 0:
			self.excel_filename = bufstr
		elif operation == 1:
			self.output_filename = bufstr
		elif operation == 2:
			bufstr = bufstr[bufstr.find("=")].upper()
			if bufstr == "TRUE":
				self.output_all = True
			else:
				self.output_all = False
		elif operation == 3:
			name = bufstr
			self.worksheetList.append(name)
		elif operation == 4:
			colonPos = bufstr.find(":")
			name = bufstr[:colonPos]
			hourVal = bufstr[colonPos+1:]
			self.specialCellsDict[name] = hourVal
		elif operation == 5:
			self.ignoreList.append(bufstr)


	def readConfigs(self):
		fconf = open(self.configFile, "r")
		conflines = fconf.readlines()
		operation = 0

		for line in conflines:
			if line[0] == "#":
				continue

			if line[0] == "\n":
				operation += 1
				continue

			self.setValues(line, operation)


if __name__ == "__main__":
	ccp = CConfigParser("xlscan.conf")
	ccp.readConfigs()



