# This module provide  functions related to create objects into the db
import sqlite3

class DbManagement:
	
	def createAlarmsTable(self, database = 'newOne.db'):
		conn = sqlite3.connect(database)
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

	def createProjectFamilyTable(database = 'newOne.db'):
		conn = sqlite3.connect(database)
		cursor = conn.cursor()
		# Drop the table if exists
		cursor.execute('''
			DROP TABLE IF EXISTS project_family
			''')    
		# Create the alarms table
		cursor.execute('''
			CREATE TABLE project_family (
				id INTEGER PRIMARY KEY AUTOINCREMENT,           
				project VARCHAR(50), 
				family VARCHAR(50)           
			)
			''')
		# Commit the changes and close the connection
		conn.commit()
		conn.close()    

	def creatcmdb_liteTable(database = 'newOne.db'):
		conn = sqlite3.connect(database)
		cursor = conn.cursor()
		# Drop the table if exists
		cursor.execute('''
			DROP TABLE IF EXISTS cmdb_lite
			''')    
		# Create the alarms table
		cursor.execute('''
			CREATE TABLE cmdb_lite (
				id INTEGER PRIMARY KEY AUTOINCREMENT, 
				server_name UNIQUE VARCHAR(50), 
				dns_name VARCHAR(50),    
				addresslist VARCHAR(50), 
				os VARCHAR(50),    
				environment VARCHAR(50), 
				family VARCHAR(50), 
				app VARCHAR(50) 
			)
			''')
		cursor.execute('''
			CREATE UNIQUE INDEX idx_cmdb_lite_server_name ON cmdb_lite (server_name);  
			''')
		cursor.execute('''
			CREATE UNIQUE INDEX idx_cmdb_lite_dns_name ON cmdb_lite (dns_name);  
			''')            
		# Commit the changes and close the connection
		conn.commit()
		conn.close() 