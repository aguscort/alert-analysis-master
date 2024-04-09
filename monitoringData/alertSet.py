import win32com.client as wcli
import os, sys, pickle, sqlite3
from datetime import datetime, timedelta
import numpy as np
import pandas as pd
import logging
import logging.config

class AlertSet:
	# Field
	__alerFieldLables = 'start_date start_time start_time_notified \
							end_date end_time_email end_time_notified \
							alert_id alert_type incident_description severity \
							server os project family admin_called \
							incident false_alert tool'.split()

	__currentAttempt = {}
	__currentAttemptTemp = {}

	 # Locations		
	__pathScript = ''
	__dbPath = ''
	__picklePath = ''
	__pathLog = ''

	# File Names
	__dbFIleName = ''
	__pickleFileName = ''

	# Alerts List
	__alerts = []

	# Parameters related to Outlook
	__mailboxes = ()
	__mailboxesAnwsers = ()

	# Date interval 
	__startDate = ''
	__endDate = ''
  
	# Parameters related to Excel
	__workbookName = ''
	__sheetName = ''
	__configSheetName = ''

	# Detailed or not:
	# Just for debugging purposes, it returns all the alerts, each column 
	# corresponding a mail, without summarizing info
	__rawData = False

	def	__registerCurrentAttempt(self, source):
		self.__currentAttempt = {'The time this query was performed was: {}' : datetime.strftime(datetime.now(), '%d/%m/%Y %H:%M')}
		self.__currentAttempt.update({'The data source used: {}' : source})		
		if source == 'outlook':
			self.__currentAttempt.update({'The mailboxes/acounts checked were: {}' : self.__mailboxes})
			self.__currentAttempt.update({'The mailboxes/acounts related to Ops mails and Admin answers were: {}' : self.__mailboxesAnwsers})
			self.__currentAttempt.update({'The first Received Date considered is: {}.' : self.__startDate})
			self.__currentAttempt.update({'The last Received Date considered is: {}.'  : self.__endDate})							
		elif source == 'file':
			if len(self.__alerts) > 0:				
				self.__currentAttempt.update({'The path where the file is located is: {}' : self.__picklePath})
				self.__currentAttempt.update({'The file name where the data is stored is: {}' :  self.__pickleFileName})								
		elif source == 'database':
			self.__currentAttempt.update({'The path were the database file is: {}' : self.__dbPath})
			self.__currentAttempt.update({'The file name where the database is located is: {}' :  self.__dbFIleName})
			self.__currentAttempt.update({'The upper limit in start_date field queried is: {}.' : self.__startDate})
			self.__currentAttempt.update({'The lower limit in start_date field queried is: {}.'  : self.__endDate})							
		else:
			self.__currentAttempt.update({'The source used was:{}' : 'None'})		

		self.__currentAttempt.update(self.__currentAttemptTemp)
		self.__currentAttemptTemp = {}

	# Functions related to Outlook importing process
	def	__organizeItemsIntoAlert(self):
		'''This function sorts the data within the pickle file and DB, if needed'''
		sortedAlerts = []
		for alert in self.__alerts:
		#Include a check to find if it's already sorted !!!! [WARNING]
			sortedAlert = {}
			for label in self.__alerFieldLables:
				sortedAlert.update({label : alert[label]})
			sortedAlerts.append(sortedAlert)
		
		self.__alerts = sortedAlerts		

	def __normalizeMails (self):
		'''Take the raw data on which each mail generate one row and summarize the
		info of those related to the same ID.'''
		dfRawAlerts = pd.DataFrame(self.__alerts)
		alertsId = set(dfRawAlerts['alert_id'])

		alertsProcessed = []
		dateFields = ('end_date','start_date')
		timeFields = ('end_time_email','end_time_notified','start_time', 'start_time_notified')
		count = 1

		# Get each group of alerts with the same ID 
		# and trransform into an only one row with all the data.
		result = {}
		for alertId in alertsId:
			dfSameId =  dfRawAlerts[dfRawAlerts['alert_id'] == alertId]
			# Get each group of alerts with the same ID and 
			# transform it into a single row with all the data.
			currentAlertInProcess = {}

			# Loop through all the columns 
			for i in range(len(dfSameId.columns)):
				# Discard the items duplicated in the column
				a = dfSameId.iloc[:,i].drop_duplicates()
				fieldValue = ''
				for index, value in a.iteritems():
					if fieldValue == '': # first Value
						fieldValue = value
					elif fieldValue == None and value != None: # new value has content
						fieldValue = value                
					elif fieldValue != None and value != None:  # both values have content          
						if a.name in dateFields or a.name in timeFields:
							if value <= fieldValue:
								fieldValue = value
					nextPair = {a.name : fieldValue} 
				currentAlertInProcess.update(nextPair)
			alertsProcessed.append(currentAlertInProcess)
			self.logger.info ('Process completed up to: {:0.2f}%'.format((100*count)/len(alertsId)))
			count += 1        

		#dfTreatedAlerts = pd.DataFrame(alertsProcessed)
		#treatedAlerts = dfTreatedAlerts.to_dict(orient='records')
		self.__currentAttemptTemp.update({'{} mails were processed.' : len(self.__alerts)})
		self.__currentAttemptTemp.update({'{} alerts registered' : len(alertsProcessed)})
		self.logger.info ('The previous {} mails were summarized into {} alerts.'.format(str(len(self.__alerts)),str(len(alertsProcessed))))
		self.__alerts = alertsProcessed
		self.__organizeItemsIntoAlert()
		
	def __processOpsMessageRawMode(self, message, tool):
		'''Process the human response mails just taking the proper values
		and assign them to the proper fields.'''
		messageProcessed = {}
		
		# We select the mails with "ID#XXXXXX" format
		# Discard when no ID# is found at the very beginning of the subject
		if message.subject.find ('ID#') != -1: 
			subject = message.subject.split(' ')
			for i in subject:
				if i.find ('ID#') != -1:
					alertId = ({'alert_id' : i[3:]})
					break

			if message.subject.upper().find('OPEN') != -1:
				timeData = {'start_date' : message.ReceivedTime.strftime('%d/%m/%Y'),
						'start_time' : None,
						'start_time_notified' : message.ReceivedTime.strftime('%H:%M:%S'),
						'end_date' : None,
						'end_time_email' : None,
						'end_time_notified' : None}
			elif message.subject.upper().find('CLOSED') != -1:  
				timeData = {'end_date' : message.ReceivedTime.strftime('%d/%m/%Y'),
						'end_time_email' : None,
						'end_time_notified' : message.ReceivedTime.strftime('%H:%M:%S'),
						'start_date' : None,
						'start_time' : None,
						'start_time_notified' : None}
			else:                        
				timeData = {'end_date' : None,
						'end_time_email' : None,
						'end_time_notified' : None,
						'start_date' : None,
						'start_time' : None,
						'start_time_notified' : None}

			# Fulfill the rest of data 
			body = message.body.split('\n')
			extraData = {}
			extraData.update({'server' : None})
			extraData.update({'alert_type' : None})
			extraData.update({'incident_description' : None})
			extraData.update({'severity' : None})
					
			extraData.update({'os' : None})             # Necesary?
			extraData.update({'project' : None})        # Necesary?
			extraData.update({'family' : None})         # Necesary?
			extraData.update({'admin_called' : 'NO'})   # Necesary?        
			extraData.update({'false_alert' : 'NO'})    # Necesary?
			extraData.update({'tool' : tool})             

			if message.body.find ('INC') != -1: 
				to = message.body.find ('INC')                    
				extraData.update({'incident' : message.body[to:to+16].translate( { ord(c):None for c in ' \n\t\r' })})
			else:            
				extraData.update({'incident' : None})
					
			# Mount the response with data
			processedMessage = {**alertId, **timeData, **extraData}
		else: 
			# Mount the empty response when the alert has no ID
			processedMessage = None            

		return processedMessage
	
	def __processMessageRawMode(self, message, tool):
		'''Process the alert mails sent by the monitoring tools just taking 
		the target values and assign them to the proper fields.'''
		messageProcessed = {}
		
		# We select the mails with "ID#XXXXXX" format
		subject = message.subject.split(' ')
		# Discard when no ID# is found at the very beginning of the subject
		if subject[0].find ('ID#') != -1:              
			alertId = ({'alert_id' : subject[0][3:]})

			if subject[1].upper().find('OPEN') != -1:
				timeData = {'start_date' : message.ReceivedTime.strftime('%d/%m/%Y'),
						'start_time' : message.ReceivedTime.strftime('%H:%M:%S'),
						'start_time_notified' : None,
						'end_date' : None,
						'end_time_email' : None,
						'end_time_notified' : None}
			elif subject[1].upper().find('CLOSED') != -1:  
				timeData = {'end_date' : message.ReceivedTime.strftime('%d/%m/%Y'),
						'end_time_email' : message.ReceivedTime.strftime('%H:%M:%S'),
						'end_time_notified' : None,
						'start_date' : None,
						'start_time' : None,
						'start_time_notified' : None}                         

			# Fulfill the rest of data 
			body = message.body.split('\n')
			extraData = {}
			for line in body:
				if line.find ('Device: ') != -1:
					extraData.update({'server' : line.split(':')[1][1:-1].strip()})
				elif not 'server' in extraData:
					extraData.update({'server' : None})
					
				if line.find ('Monitor Type: ') != -1:                
					extraData.update({'alert_type' : line.split(':')[1][1:-1]})                
				elif not 'alert_type' in extraData:
					extraData.update({'alert_type' : None})
					
				if line.find ('Instance: ') != -1:                
					#Workaround to process windows units properly                                
					extraData.update({'incident_description' : line.replace(': ', '#').split('#')[1][:-1].strip()})
				elif not 'incident_description' in extraData:
					extraData.update({'incident_description' : None})
					
				if line.find ('Severity: ') != -1:                
					extraData.update({'severity' : line.split(':')[1][1:-1].upper()})                
				elif not 'severity' in extraData:
					extraData.update({'severity' : None})
					
				extraData.update({'os' : None})             # Necesary?
				extraData.update({'project' : None})        # Necesary?
				extraData.update({'family' : None})         # Necesary?
				extraData.update({'admin_called' : 'NO'})   # Necesary?
				extraData.update({'incident' : None})
				extraData.update({'false_alert' : 'NO'})    # Necesary?
				extraData.update({'tool' : tool})                                                    
					
			# Mount the response with data
			processedMessage = {**alertId, **timeData, **extraData}
		else: 
			# Mount the empty response when the alert has no ID
			processedMessage = None            

		return processedMessage
		
	def __storeMessages(self, messages, tool, fromTool = True):
		'''Sotre each message extracted from the mailboxes in a list
		once has been processed.'''
		for message in messages:
			# Process each message
			if fromTool == True:
				messageCool = self.__processMessageRawMode(message, tool)
			else:
				messageCool = self.__processOpsMessageRawMode(message, tool)                   
			if messageCool != None:
				self.__alerts.append(messageCool)
	
	def __extracAlarmsFromOutlook(self):
		'''Revisa los mailboxes indicados para extraer info de las alertas'''
		alerts = []
		outlook = wcli.Dispatch("Outlook.Application").GetNamespace("MAPI")
		self.logger.info ('Connecting to Outlook App.') 
		
		for oAccount in outlook.Session.Accounts:                              
			for m in self.__mailboxes:
					if oAccount.SmtpAddress == m['account']:
						for mailbox in m['mailboxes']:
							store = oAccount.DeliveryStore
							inbox = store.GetDefaultFolder(6).Folders(mailbox)
							messages = inbox.Items.Restrict("[ReceivedTime] >= '" + self.__startDate +  "' \
									AND [ReceivedTime] <= '" + self.__endDate + "'")
							self.logger.info ('Checking ' + mailbox + '...')
							self.__storeMessages(messages, m['tool'], True)
							self.logger.info ('Recovered ' + str(len(self.__alerts)) + ' mails from mailboxes so far...')


		
		# Second round: processoiing Ops notifications
		for oAccount in outlook.Session.Accounts:                              
			for m in self.__mailboxesAnwsers:
					if oAccount.SmtpAddress == m['account']:
						for mailbox in m['mailboxes']:
							store = oAccount.DeliveryStore
							if type(mailbox) ==  tuple:
								inbox = store.GetDefaultFolder(6).Parent.Folders(mailbox[0]).Folders(mailbox[1])
							else:    
								inbox = store.GetDefaultFolder(6).Parent.Folders(mailbox)
							messages = inbox.Items.Restrict("[ReceivedTime] >= '" + self.__startDate +  "' \
										AND [ReceivedTime] <= '" + self.__endDate + "'")
							self.logger.info ('Checking ' + mailbox + '...') if type(mailbox) !=  tuple else self.logger.info ('Checking ' + mailbox[1] + '...')
							mailbox
							self.__storeMessages(messages, m['tool'], False)
							self.logger.info ('Recovered ' + str(len(self.__alerts)) + ' mails from mailboxes so far...')

		outlook = None        

		if self.__rawData == True:
			self.logger.info ('Recovered {} mails without summarizing data.'.format(str(len(alerts))))
		else:
			self.__normalizeMails()
		self.logger.info ("Extracted mails from {} to {}".format(self.__startDate, self.__endDate))


	# These functions provide the feauture to serlize the list alerts into a picke file
	def __setHistData(self, hFile):
		'''This function stores the data within a list into a pickle file'''
		try:
			with open(hFile, 'wb') as f:
					pickle.dump(self.__alerts, f)
		except Exception as e: 
			self.logger.warning ("While saving data into a file, an exception of type {0} occurred. Arguments:\n{1!r}".format(type(e).__name__, e.args))

	def __getHistData(self, hFile, getInterval=False):
		'''This functon populate a list with the data already saved in to a pickle file'''
		try:
			if not os.path.isfile(hFile):
				print('No file was found')
				self.logger.warning ('File {} was not found.'.format(hFile))
				return ()

			with open(hFile, 'rb') as f:
					self.__alerts = pickle.load(f)

			dfTemp = pd.DataFrame(self.__alerts)			
			self.__currentAttemptTemp.update({'The first date of an alert saved within the file is: {}.'  : dfTemp['start_date'].min()})
			self.__currentAttemptTemp.update({'The last date of an alert saved within the file is: {}.'  : dfTemp['start_date'].max()})							

			# When the data interval must be applied within the file
			if getInterval == True:	
				# Set the date as index in order to filter later
				dfTemp['start_date_ix'] = pd.to_datetime(dfTemp['start_date'])
				dfTemp.set_index('start_date_ix', inplace=True)		
				# Filter the dataframe to the date interval 
				# then we transform it into a dictionary with records parameter
				# so each row becomes a dictionary where key is column name 
				# and value is the data in the cell 		
				self.__alerts = dfTemp.loc[str(datetime.strptime(self.__startDate, '%d/%m/%Y').date()):str(datetime.strptime(self.__endDate, '%d/%m/%Y').date())].to_dict('records')
				self.__currentAttemptTemp.update({'The number of alerts within the file returned between {}'  : self.__startDate})
				self.__currentAttemptTemp.update({'                and {}'  : self.__endDate})							
				self.__currentAttemptTemp.update({'                is: {}.'  : len(self.__alerts)})							
				
			self.__organizeItemsIntoAlert()

		except Exception as e: 
			self.logger.warning ("While retrieving data from a file, an exception of type {0} occurred. Arguments:\n{1!r}".format(type(e).__name__, e.args))


	# These functions related to storing and retrieving info into and from a sqlite database 
	def __convertDbAlertsToDict(self):
		'''This method transforms the list returned by the query into a dict'''
		sortedAlerts = []
		for alert in self.__alerts:
			sortedAlert = {}
			# The values are already ordered so we just need to add the key
			i = 0
			for label in self.__alerFieldLables:
				sortedAlert.update({label : alert[i]})
				i += 1
			sortedAlerts.append(sortedAlert)
		
		self.__alerts = sortedAlerts

	def __createAlarmsTable(self):
		try:
			conn = sqlite3.connect(os.path.join(self.__dbPath, self.__dbFIleName))
			cursor = conn.cursor()
			# Drop the table if exists
			cursor.execute('''
				DROP TABLE IF EXISTS alarms
				''')    
			# Create the alarms table
			cursor.execute('''
				CREATE TABLE alarms (
					start_date DATETIME, 
					start_time DATETIME, 
					start_time_notified DATETIME, 
					end_date DATETIME, 
					end_time_email DATETIME, 
					end_time_notified DATETIME,
					alert_id VARCHAR(10) PRIMARY KEY, 
					alert_type VARCHAR(20), 
					incident_description VARCHAR(100), 
					severity VARCHAR(20), 
					server VARCHAR(50), 
					os VARCHAR(20), 
					project VARCHAR(50), 
					family VARCHAR(50), 
					admin_called BOOLEAN, 
					incident VARCHAR(20), 
					false_alert BOOLEAN, 
					tool VARCHAR(20)
				)
				''')
			# Create index
			cursor.execute('''
				CREATE UNIQUE INDEX idx_alarms_alert_id ON alarms (alert_id);  
				''')

			# Commit the changes and close the connection
			conn.commit()
			conn.close()
		except Exception as e: 
			self.logger.warning ("While trying to create the alerts Table, an exception of type {0} occurred. Arguments:\n{1!r}".format(type(e).__name__, e.args))		

	def __populateAlarmsTable(self):
		'''Get a list of dictionaries were the alarms get stored and populate a table in a sqlite database'''
		count = 1
		finalOutcome = False
		toCheck = ('incident', 'start_time_notified', 'end_date', 'end_time_email', 'end_time_notified')

		conn = sqlite3.connect(os.path.join(self.__dbPath, self.__dbFIleName))
		cursor = conn.cursor()
		try:
			cursor.execute('BEGIN TRANSACTION') 
			# Do we need to do this in order to improve the Insert??
			cursor.execute('DELETE FROM alarms WHERE start_date >= ? \
							AND start_date  <= ?', (self.__startDate, self.__endDate)) 

			for alert in self.__alerts:    
				query = '''
					INSERT OR IGNORE INTO alarms (alert_id, alert_type, incident_description,
								severity, server, os, project, family, admin_called, incident,
								false_alert, start_date, start_time, start_time_notified, end_date,
								end_time_email, end_time_notified, tool) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
				'''  

				vars = (alert['alert_id'], alert['alert_type'], alert['incident_description'], alert['severity'], 
						alert['server'], '', '', '', alert['admin_called'], alert['incident'], '', alert['start_date'], 
						alert['start_time'], alert['start_time_notified'], alert['end_date'], alert['end_time_email'],
						alert['end_time_notified'], alert['tool'])    
				cursor.execute(query, vars) 

				if cursor.rowcount < 1:                    
					for i in toCheck:  
						if alert[i] != None:                      
							query = 'UPDATE alarms SET '  + i + ' = "' + alert[i] +  '" WHERE alert_id = ' + alert['alert_id']  +  ' AND ' + i + ' is Null'                            
							cursor.execute(query) 
							if cursor.rowcount < 1:
								self.logger.debug (query)
				
				self.logger.info ('Process completed up to: {:0.2f}%'.format((100*count)/len(self.__alerts)))
				count += 1
		except Exception as e: 
			cursor.execute('ROLLBACK') 
			self.logger.warning ("The error {} ocurred when trying to insert this element:".format(type(e).__name__))
			self.logger.warning (vars)
			finalOutcome = False
		else: 
			cursor.execute('COMMIT') 
			finalOutcome = True            
		conn.close()

		if finalOutcome == True:
			self.logger.info ('The process of saving the alerts into the database ended successfully')
			self.__currentAttemptTemp.update({'alerts_registered = {}' : len(self.__alerts)})				
		else:
			self.logger.warning ('An error was founded during the database populate process')
			self.__currentAttemptTemp.update({'alerts_registered = {}' : 0})
	
	def __getAlarmsByID (database = 'newOne.db', alert_id = '0'):
		'''Return all the alerts acording thir ID'''
		pass
	
	def __getAlarmsByDate (self, informed = True):
		'''Return all the alerts between two dates'''
		try:		
			conn = sqlite3.connect(os.path.join(self.__dbPath, self.__dbFIleName))
			cursor = conn.cursor()

			if informed == True:
				alerts = cursor.execute('SELECT * FROM alarms WHERE start_date >= ? AND start_date  <= ? AND start_time_notified is not null ORDER BY start_date, start_time', (self.__startDate, self.__endDate)).fetchall()    
			else:        
				alerts = cursor.execute('SELECT * FROM alarms WHERE start_date >= ? AND start_date  <= ? ORDER BY start_date, start_time', (self.__startDate, self.__endDate)).fetchall()    
			self.logger.info ("{} rows returned by the query.".format(len(alerts)))

			conn.close()   
			self.__alerts = alerts 
			self.__convertDbAlertsToDict()
			self.__currentAttemptTemp.update({'The number of rows returned by the query was: {}' : len(alerts)})

			update({'alerts_retrieved' : len(alerts)})		
		except Exception as e: 
			self.logger.warning ("An exception of type {0} occurred. Arguments:\n{1!r}".format(type(e).__name__, e.args))
			self.__currentAttemptTemp.update({'Alerts retrieved = {}' : 0})		

	def __removeAlarmsByDate (self):
		'''Remove the alarms between two dates'''
		pass
	
	def __emptyAlarmsTable(self):
		'''Remove all elements from alarms table'''
		try:
			conn = sqlite3.connect(os.path.join(self.__dbPath, self.__dbFIleName))
			cursor = conn.cursor()
			# Remove all elements from alarms table
			cursor.execute('''
				DELETE FROM alarms
				''')
			# Commit the changes and close the connection
			conn.commit()
			conn.close()  
		except Exception as e: 
			self.logger.warning ("An exception of type {0} occurred. Arguments:\n{1!r}".format(type(e).__name__, e.args))	

	# These functions related to writting info into Excel doc 
	def __writeToExcel (self):
		'''Write the list of alarmas into Excel'''
		wsFound = False
		excel = wcli.Dispatch('Excel.Application')
	  
		for wkb in excel.Workbooks:
			if wkb.Name == self.__workbookName:
				#excel.Visible = False     
				wsFound = True
				ws = wkb.Worksheets(self.__sheetName)
				firstRow = ws.Range("H" + str(ws.Rows.Count)).End(-4162).Row + 1
				row = firstRow
				for alarm in self.__alerts:
					column = 1            
					for item in alarm: 
						#print (alarm[item])                     
						if column == 7:
							column += 1
						while ws.Columns(column).Hidden == True:
							column += 1
						# Just until "L" column
						if column <= 12 or column == 17:
							if alarm[item] != None:
								# Trick to avoid the mm/dd dd/mm issue... workaround
								if column == 1 or column ==4:                            
									ws.Cells(row, column).NumberFormat = 'dd/mm/yyyy;@'
									ws.Cells(row, column).Value = str(datetime.strptime(alarm[item], '%d/%m/%Y').date() + timedelta(days=13))
									ws.Cells(row, column).Value = ws.Cells(row, column).Value - timedelta(days=13)
									pass
								else:
									ws.Cells(row, column).Value = alarm[item]                  
						else:
							pass   
						column += 1                        
					ws.Range("M" + str(row)).Formula = '=IFERROR(VLOOKUP(L'+ str(row) +' & "*",Config!$N$2:$Q$9855,4,FALSE),VLOOKUP(L'+ str(row) +' & "*",Config!$O$2:$Q$9855,3,FALSE))'
					ws.Range("O" + str(row)).Formula = '=VLOOKUP(N'+ str(row) +',Config!H$1:I$150,2,FALSE)'
					ws.Range("S" + str(row)).Formula = '=TRIM(C'+ str(row) +')-TRIM(B'+ str(row) +')'
					ws.Range("T" + str(row)).Formula = '=IF(OR(ISBLANK(A'+ str(row) +')=TRUE,ISBLANK(B'+ str(row) +')),"N/A",IF(OR(B'+ str(row) +'<0.291666666,B'+ str(row) +'>0.833333333),"Night",IF(WEEKDAY(A'+ str(row) +',2)>5,"Weekend",IF(COUNTIF(Config!$L$2:$L$25,A'+ str(row) +')>0,"Bank Holiday","Day"))))'
					row += 1
					self.logger.debug ("Row {} written".format(row -1))
		excel.Visible = True            
		excel = None
		
		if wsFound == False:
			self.logger.warning("The Workbook is not currently open")
		else:
			self.logger.info ("Alarms were written from row {} to {} row (A total of {}).".format(firstRow, row - 1, row - firstRow + 1))        
		self.logger.info('Process ended')


	# Constructor and desctructor
	def __init__(self, startDate = '01/01/1971', endDate = '01/01/1971'):
		self.__startDate = startDate
		self.__endDate = endDate
		logFileName  = os.path.join(self.__pathLog , 'alertSet.log')
		print ('d:\\alertSet.log')
		# Get logging		
		self.logger = logging.getLogger(__name__)		
		logging.config.dictConfig({
			'version': 1,
			'disable_existing_loggers': False,  # this fixes the problem
			'formatters': {
				'standard': {
					'format': '%(asctime)s [%(levelname)s] %(name)s: %(message)s'
				},
			},
			'handlers': {
				'default': {
					'level':'DEBUG',
					'class': 'logging.handlers.RotatingFileHandler',
					'filename': 'd:\\alertSet.log', # logFileName,
					'formatter': 'standard',
					'maxBytes': 10485760, # 10MB
					'backupCount': '20',
					'encoding': 'utf8'		            
				},		       
			},
			'loggers': {
				'': {
					'handlers': ['default'],
					'level': 'INFO',
					'propagate': True
				}
			}
		})

		self.logger.info('New alertSet object created')

	def __del__(self):
		self.logger.info('alertSet object destroyed')


	# AlertSet related methods		
	def retrieveAlerts(self, source, setInterval=False):
		'''Retrieve an alert set from diferent sources'''
		if source == 'outlook':
			self.__extracAlarmsFromOutlook()
		elif source == 'file':
			self.__getHistData(os.path.join(self.__picklePath, self.__pickleFileName), setInterval)
		elif source == 'database':
			self.__getAlarmsByDate (False)
		else:
			print ('No source was identified: try "outlook", "file" or "database"')	
		self.__registerCurrentAttempt(source)

	def saveAlerts(self, source):
		'''Save an alert set to diferent media'''
		if source == 'file':
			self.__alerts = self.__setHistData(os.path.join(self.__picklePath, self.__pickleFileName))
		elif source == 'database':
			self.__populateAlarmsTable()	
		elif source == 'excel':				
			self.__writeToExcel()
		else:
			print ('No media was identified: try "file", "excel" or "database"')	

	def append(self, a):
		'''Append new alerts to a set already formed'''
		self.alerts.append(a)
		pass

	def showAlerts(self, start=1, end=5): #len(self.__alerts)):
		'''Print a couple of alerts from the current set'''
		for i in self.__alerts[start:end]:
			print(i)	

	def getAlerts(self):
		'''Return the current alert set'''
		return self.__alerts	

	def showCurrentAttempt(self):
		for i in self.__currentAttempt:
			print(i.format(self.__currentAttempt[i]))

	# Date related properties
	def getStartDate(self):
		return self.__startDate

	def getEndDate(self):		 
		return self.__endDate

	def setStartDate(self, startDate):
		self.__startDate = startDate

	def setEndDate(self, endDate):		 
		self.__endDate = endDate

	# Mailboxes related properties
	def getMailboxes(self):		 
		return self.__mailboxes
	
	def setMailboxes(self, mailboxes):
		try:
			self.__mailboxes = tuple(mailboxes)
		except:
			self.logger.warning ('Error: no mailboxes were added')

	def getMailboxesAnwsers(self):		 
		return self.__mailboxesAnwsers
	
	def setMailboxesAnwsers(self, mailboxesAnwsers):
		try:
			self.__mailboxesAnwsers = tuple(mailboxesAnwsers)
		except:
			self.logger.warning ('Error: no mailboxes were added')

	# Database location related methods
	def getDbPath(self):
		return self.__dbPath

	def setDbPath(self, dbPath):
		self.__dbPath = dbPath

	def getDbFileName(self):		 
		return self.__dbFIleName

	def setDbFileName(self, dbFIleName):		 
		self.__dbFIleName = dbFIleName

	# Serialize file location related methods
	def getPickleFileName(self):
		return self.__pickleFileName

	def setPickleFileName(self, pickleFileName):
		self.__pickleFileName = pickleFileName

	def getPicklePath(self):		 
		return self.__picklePath

	def setPicklePath(self, picklePath):		 
		self.__picklePath = picklePath

	# Excel file related methods
	def getWorkbookName(self):
		return self.__workbookName

	def setWorkbookName(self, workbookName):
		self.__workbookName = workbookName

	def getSheetName(self):		 
		return self.__sheetName

	def setSheetName(self, sheetName):		 
		self.__sheetName = sheetName

	# Log location related methods
	def getPathLog(self):		 
		return self.__pathLog

	def setPathLog(self, pathLog):		 
		self.__pathLog = pathLog