import datetime
import re
import sys
from sheetdata import CSheetData

sd = CSheetData() #access excell worksheet hardcoded data


class CEmployee:
	name = ""
	scheduled_hours = 0
	overtime_scan = False

	def __init__(self):
		#mutable objects, so they can be passed "by reference"
		self.guarding_hours = [0]
		self.private_hours = [0]
		self.class_hours = [0]
		self.break_hours = [0]
		self.supervisor_hours = [0]
		self.misc_hours = [0]
		self.total_hours = [0]

	#Add regex time format error checking??
	def convertTimeISO(self, tstr, mstr): 
		meridiem_regex = re.compile(r"[^amp]+")
		meridiemString = meridiem_regex.sub("", tstr)

		num_regex = re.compile(r"[^\d:]+")
		numString = num_regex.sub("", tstr)

		#for an overall meridiem, do not overwrite current meridiem
		if meridiemString == "":
			meridiemString = mstr 

		colonPos = numString.find(":")
		if colonPos == -1:
			tstr = numString + ":00" + meridiemString
		else:
			tstr = numString + meridiemString

		tempDatetime = datetime.datetime.strptime(tstr, "%I:%M%p")
		minutes = tempDatetime.minute
		if minutes < 10:
			minutes = "0" + str(minutes)

		tstr = str(tempDatetime.hour) + ":" + str(minutes)
		return tstr


	def getTimeMeridiem(self, timeString):
		meridiem = "am"
		isAM = isPM = False

		if "AM" in timeString.upper():
			isAM = True
		if "PM" in timeString.upper():
			isPM = True

		if isAM and not isPM:
			meridiem = "am"
		elif isPM and not isAM:
			meridiem = "pm"

		return meridiem


	def handleAbsoluteTime(self, timeString):
		options = {"0": 0.0, "15": .25, "30": .5, "45": .75, "60": 1.0, "1": 1.0}
		abs_time_regex = re.compile(r"[^\d]+")

		timeString = abs_time_regex.sub("", timeString)
		hourAmount = options.get(timeString, .5)  #if not in lookup, return .5

		return hourAmount


	def handleRangedTime(self, timeString):
		hyphenPos = timeString.find("-")
		meridiem = self.getTimeMeridiem(timeString)

		tstart = timeString[:hyphenPos]
		tstart = self.convertTimeISO(tstart, meridiem)
		startTime = datetime.datetime.strptime(tstart, "%H:%M")

		tend = timeString[hyphenPos+1:]
		tend = self.convertTimeISO(tend, meridiem)
		endTime = datetime.datetime.strptime(tend, "%H:%M")

		if startTime > endTime:
			startTime -= datetime.timedelta(hours=12)

		deltaTime = endTime - startTime
		hourAmount = deltaTime.seconds / 3600.0  #floting point division

		return hourAmount


	def handleMultipleRangedTimes(self, timeString):
		timeList = timeString.split("_")
		hourAmount = 0

		for tstr in timeList:
			hourAmount += self.handleRangedTime(tstr)

		return hourAmount


	def parseTimeString(self, timeString):
		hourAmount = .5
		timeString = timeString.replace(" ", "_")
		time_string_regex = re.compile(r"[a-zA-Z\s]*[^\damp_:-]+")

		timeString = time_string_regex.sub("", timeString)
		hyphenPos = timeString.find("-")
		underPos = timeString.find("_")

		#If no hyphen, it is an absolute time. (e.g. 45 mins)
		if hyphenPos == -1:
			hourAmount = self.handleAbsoluteTime(timeString)
		elif underPos == -1:
			hourAmount = self.handleRangedTime(timeString)
		else:
			hourAmount = self.handleMultipleRangedTimes(timeString)

		return hourAmount


	def getHourAmount(self, cellValue):
		hourAmount = .5 #default
		if cellValue == None:
			return hourAmount

		#Look up hour amount for special time slots
		hourAmount = sd.getSpecialHourAmount(cellValue)
		if hourAmount > -1:
			return hourAmount

		#Lookup hour amount for a group lesson
		hourAmount = sd.getGroupLessonHourAmount(cellValue)
		if hourAmount > -1:
			return hourAmount

		#handle time ranges and absolute times
		timeString = cellValue.strip()
		hourAmount = self.parseTimeString(timeString)
		
		return hourAmount


	def setAttributes(self, cellColor, cellValue, cellCol):
		if type(cellValue) == datetime.time:
			cellValue = datetime.time.strftime(cellValue, "%H:%M:%S")	

		if cellValue is None:
			if cellColor == sd.blankColor0 or cellColor == sd.blankColorF:
				return 3
		else:
			for istr in sd.ignoreList: #check for value in ignore list
				if istr in cellValue.replace(" ", "").upper():
					return 4

		if cellColor == sd.headerColor:
			return 5

		if cellCol == "A":
			self.name = cellValue
			return 1
		elif cellCol == "B":
			self.scheduled_hours = self.getHourAmount(cellValue)
			return 2

		hourAmount = self.getHourAmount(cellValue)
		self.total_hours[0] += hourAmount

		employeeAttribute = sd.getAttrObjFromColor(cellColor, self)
		employeeAttribute[0] += hourAmount

		return 0


	def mergeHoursWith(self, newEmployee):
		self.scheduled_hours += newEmployee.scheduled_hours
		self.guarding_hours[0] += newEmployee.guarding_hours[0]
		self.private_hours[0] += newEmployee.private_hours[0]
		self.class_hours[0] += newEmployee.class_hours[0]
		self.supervisor_hours[0] += newEmployee.supervisor_hours[0]
		self.break_hours[0] += newEmployee.misc_hours[0]
		self.total_hours[0] += newEmployee.total_hours[0]


	def checkOvertime(self):
		rstring = ""

		if self.scheduled_hours > 80:
			rstring = "[ OVERTIME DECTECTED ]\n"
		elif self.total_hours[0] > 80:
			rstring = "[ OVERTIME DECTECTED ]\n"

		return rstring


	def checkTotalHours(self):
		calculatedHours = self.total_hours[0]
		scheduledHours = self.scheduled_hours
		rstring = ""

		if scheduledHours != calculatedHours and not self.overtime_scan:
			rstring = "[Warning:] Time discrepancy found! Hours scheduled does not match total hours.\n"
			rstring += "Calculated hours: " + str(calculatedHours) + "\n"
			rstring += "Scheduled hours: " + str(scheduledHours) + "\n"
			rstring += "\n"

		return rstring


	def toString(self):
		rstring = "Name: " + self.name + "\n"
		rstring += "Scheduled Hours: " + str(self.scheduled_hours) + "\n"

		if not self.overtime_scan:
			rstring += "------ Hours Breakdown ------\n"
			rstring += "Guarding: " + str(self.guarding_hours[0]) + "\n"
			rstring += "Private Lessons: " + str(self.private_hours[0]) + "\n"
			rstring += "Group Lessons: " + str(self.class_hours[0]) + "\n"
			rstring += "Deck Supervising: " + str(self.supervisor_hours[0]) + "\n"
			rstring += "Breaks: " + str(self.break_hours[0]) + "\n"
			rstring += "Miscellaneous: " + str(self.misc_hours[0]) + "\n"
			rstring += "Total: " + str(self.total_hours[0]) + "\n"

		rstring += "\n"
		return rstring


