class Alert:
	__alerFieldLables = 'start_date start_time start_time_notified \
							end_date end_time_email end_time_notified \
							alert_id alert_type incident_description severity \
							server os project family admin_called \
							incident false_alert tool'.split()
	__alert_id = ''
	__alert_type = ''
	__incident_description = ''
	__severity = ''
	__server = ''
	__os = ''
	__project = ''
	__family = ''
	__admin_called = ''
	__incident = ''
	__false_alert = ''
	__start_date = '' 
	__start_time = ''
	__start_time_notified = ''
	__end_date = ''    
    __end_time_email = ''
    __end_time_notified = ''
    __tool = ''


	def __init__(self):

	def __del__(self):
		pass

	def getAlert(self):
		pass
	
	def setAlert(self, alert={}):
		self.__alert_id = alert['alert_id']
		self.__alert_type = alert['']
		self.__incident_description = alert['']
		self.__severity = alert['']
		self.__server = alert['']
		self.__os = alert['']
		self.__project = alert['']
		self.__family = alert['']
		self.__admin_called = alert['']
		self.__incident = alert['']
		self.__false_alert = alert['']
		self.__start_date = alert['']
		self.__start_time = alert['']
		self.__start_time_notified = alert['']
		self.__end_date = alert['']    
		self.__end_time_email = alert['']
		self.__end_time_notified = alert['']
		self.__tool = ''
