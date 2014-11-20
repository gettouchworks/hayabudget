import sqlite3



class DB(object):

	def __init__(self, dbname="test.db"):
		print "open db"
		self.conn = sqlite3.connect(dbname)
		self.conn.row_factory = sqlite3.Row
		self.cursor = self.conn.cursor()

	def getData(self, sql):
		self.cursor.execute(sql)
		return self.cursor.fetchall()

	def getLine(self, sql):
		self.cursor.execute(sql)
		return self.cursor.fetchone()
		if rs:
			return rs[0]
		return None

	def runSql(self, sql):
		self.cursor.execute(sql)
		self.conn.commit()
	
	def closeDb(self):
		self.cursor.close()
		self.conn.close()

	def __del__(self):
		try:
			# print "close db"
			self.cursor.close()
			self.conn.close()
			pass
		except:
			pass

def prate(money):
	rmoney = money-3500
	rate = [0.03, 0.1, 0.2, 0.25, 0.3, 0.35, 0.45]
	dive = [0, 105, 555, 1005, 2755, 5505, 13505]
	return max([rmoney*x[0]-x[1] for x in zip(rate,dive)]+[0])

if __name__ == "__main__":
	for i in range(100, 3000):
		print i*100, i*100-prate(i*100)


	
	# db = DB()