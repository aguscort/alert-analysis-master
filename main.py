from monitoringData.alertSet import *
import numpy as np
import pandas as pd

pathScript = os.path.join('D:','repos','alert-analysis')


asMensual = AlertSet((datetime.now() - timedelta(days=3)).strftime('%d/%m/%Y'), (datetime.now() + timedelta(days=1)).strftime("%d/%m/%Y")  )

asMensual.setMailboxes(({'tool' : 'proactivenet', 
				'account' : 'ITOIS_Operations@euipo.europa.eu',
				'mailboxes' : ['ALARMS_WEBSITE', 'MINOR_ALARMS']},
				{'tool' : 'truesight', 
				'account' : 'MonitoringAlerts@euipo.europa.eu',
				'mailboxes' : ['_ALARMS_ALL']}))
asMensual.setMailboxesAnwsers(({'tool' : 'proactivenet', 
				'account' : 'ITOIS_Operations@euipo.europa.eu',
				'mailboxes' : [('IP - ALARMS', 'OWS Alarms')]},
				{'tool' : 'truesight', # This is not true anymore
				'account' : 'MonitoringAlerts@euipo.europa.eu',
				'mailboxes' : ['_ALARMS_REPORTED']}))

# Locations
asMensual.setDbPath (os.path.join(pathScript , 'db'))
asMensual.setDbFileName('alertsFromMail.db')
asMensual.setPicklePath (os.path.join(pathScript , 'data'))
asMensual.setPickleFileName('currentAlerts.pckl')
asMensual.setPathLog (pathScript)

# Parameters related to Excel
asMensual.setWorkbookName ('Monthly Alerts Report 2018.xlsx')
asMensual.setSheetName ('Alarms Reported')


#print('Outlook===================')
#asMensual.retrieveAlerts('outlook')
#asMensual.showAlerts()
#asMensual.saveAdf = lerts('database')
#print('DB===================')
#asMensual.retrieveAlerts('database')
#asMelen(self.__alerts)nsual.showAlerts()
#asMensual.saveAlerts('file')


#print('DATABASE===================')
asMensual.retrieveAlerts('outlook')
asMensual.showCurrentAttempt()
#print(asMensual.getPathLog())
asMensual.saveAlerts('excel')
asMensual.saveAlerts('database')
#asMensual.saveAlerts("file")
#print('FILE===================')
##print('FILE===================')
##asMensual.retrieveAlerts('file', True)
#asMensual.showCurrentAttempt()
#asMensual.showAlerts()
#asMensual.saveAlerts('excel')

#asMensual.saveAdf = lerts('database')
#df = asMensual.getAlerts()
#df.head()