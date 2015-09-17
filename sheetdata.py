
class CSheetData:
	# Globals for timesheet colors
	guardColor = "FFFFEB9C"
	privateConflictColor = "FFFFC7CE"
	privateColor = "FF7030A0"
	campColor = "FF00B0F0"
	breakColor = "FFC6EFCE"
	classColor = "FFFABF8F"
	headerColor = "FF1F497D"
	#miscColor = "FFB7B7B7"
	deckSupervisingColor = "FFFFFF00"
	blankColor0 = "00000000"
	blankColorF = "FFFFFFFF"

	#--- Mutable data structures are shared accross instances ---# 

	ignoreList = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY",
				  "SATURDAY", "SUNDAY", "LARGEPOOL", "SMALLPOOL"]

	#Group lesson time lookup
	lessonDict = {"PIKE": .5, "EEL": .5, "RAY": .5, "SHRIMP": .5,
				  "STARFISH": .75, "POLLY": .75, "GUPPY": .75, "MINNOW": .75, 
				  "PORPOISE": .75, "FISH": .75, "FLYINGFISH": .75, "SHARK": .75, 
				  "ADULT": 1}

	spValuesDict = {"WATEREX": 1.0}


	def getSpecialHourAmount(self, cellValue):
		tmpStr = cellValue.strip().replace(" ", "").upper()
		hourAmount = self.spValuesDict.get(tmpStr, -1)
		return hourAmount

	#Lookup hour amount for a group lesson
	def getGroupLessonHourAmount(self, cellValue):
		hourAmount = -1
		cellString = cellValue.strip().replace(" ", "").upper()

		if cellString in set(k.upper() for k in self.lessonDict):
			hourAmount = self.lessonDict[cellString]

		return hourAmount

	#returns an object reference to a mutable employee attribute based on cell color
	def getAttrObjFromColor(self, cellColor, employee):
		attributeDict = {self.guardColor: employee.guarding_hours, 
						 self.privateColor: employee.private_hours, 
						 self.campColor: employee.class_hours,
						 self.breakColor: employee.break_hours,
						 self.classColor: employee.class_hours,
						 self.deckSupervisingColor: employee.supervisor_hours} 

		attrObj = attributeDict.get(cellColor, employee.misc_hours)
		return attrObj








