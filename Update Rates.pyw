import datetime
import pandas as pd
import pandas.io.sql
import pymssql
import tkinter as tk

now = datetime.datetime.now()
year = str(now.year)

def main():
	global master
	master = tk.Tk()
	background = 'azure2'
	font = "Helvetica 10 bold underline"

	#Labels for GUI
	tk.Label(master, text="Quote Number", background=background, font = font).grid(row=0)
	tk.Label(master, text="Purchased Material %", background=background, font = font).grid(row=10)
	tk.Label(master, text="Fabricated Material %", background=background, font = font).grid(row=11)
	tk.Label(master, text="Labor Cost %", background=background, font = font).grid(row=12)
	tk.Label(master, text="Burden Cost %", background=background, font = font).grid(row=13)
	tk.Label(master, text="Service Cost %", background=background, font = font).grid(row=14)

	#build GUI
	menubar = tk.Menu(master)
	menubar.add_command(label="Update", command=combinedUpdate)
	menubar.add_command(label="Recalculate", command=combinedRecalc)
	menubar.add_command(label="Reset", command=resetFromMenu)
	menubar.add_command(label="Help", command=helpDoc)
	menubar.add_command(label="Restart", command=restart)
	menubar.add_command(label="Quit", command=master.quit)
	master.bind('<Escape>', close)
	master.bind('<F2>', resetFromMenu)
	master.bind('<Control-r>', combinedRecalc)
	master.bind('<Control-u>', combinedUpdate)			
	master.bind('<Control-h>', helpDoc)
	master.bind('<Control-R>', combinedRecalc)
	master.bind('<Control-U>', combinedUpdate)			
	master.bind('<Control-H>', helpDoc)
	master.config(menu=menubar)
	master.title("Update Quote Rates")
	master.minsize(width=350, height=100)
	master.configure(background=background)

	#Entry boxes
	global e
	e = tk.Entry(master)
	e.grid(row=0, column=1, sticky="NSEW", padx=10, pady=10)
	global e2
	e2 = tk.Entry(master)
	e2.grid(row=10, column=1, sticky="NSEW", padx=2, pady=2)
	global e3
	e3 = tk.Entry(master)
	e3.grid(row=11, column=1, sticky="NSEW", padx=2, pady=2)
	global e4
	e4 = tk.Entry(master)
	e4.grid(row=12, column=1, sticky="NSEW", padx=2, pady=2)
	global e5
	e5 = tk.Entry(master)
	e5.grid(row=13, column=1, sticky="NSEW", padx=2, pady=2)
	global e6
	e6 = tk.Entry(master)
	e6.grid(row=14, column=1, sticky="NSEW", padx=2, pady=2)
	e.focus_set()

#connect to the database
conn = pymssql.connect(
    host=r"host",
    user=r"user",
    password="password",
    database="database"
)

cursor = conn.cursor()

#function to update rates
def updateRates():
	#make sure the quote number exists in the database
	verifySQL = "SELECT ID FROM QUOTE WHERE ID IN (SELECT ID FROM QUOTE WHERE ID = '" + e.get() + "')"
	v = pandas.io.sql.read_sql(verifySQL, conn)
	D = pd.DataFrame(v)
	if D.empty:
		global top
		top = tk.Toplevel(master)
		top.title('Error')
		msg = tk.Message(top, text="Please Check Your Quote Number", width=750)
		msg.grid(row=0, column=1)
		return
	#prevent 2 from being calculated as 20 by forcing all percentages to be two digits long
	elif len(e2.get()) != 2 or len(e3.get()) != 2 or len(e4.get()) != 2 or len(e5.get()) != 2 or len(e6.get()) != 2:
			top = tk.Toplevel(master)
			top.title('Error')
			msg = tk.Message(top, text="All Percentages Must Be 2 Digits Long", width=750)
			msg.grid(row=0, column=1)
			return
	try:
		top = tk.Toplevel(master)
		top.title('Updating')
		msg = tk.Message(top, text="Rates Successfully Updated         ", width=750)
		msg.grid(row=0, column=1)
		#read in all lines from quote and store them in a dataframe
		SQL = "SELECT LINE_NO FROM CR_QUOTE_LIN_PRICE WHERE QUOTE_ID = '" + e.get() + "'"
		q = pandas.io.sql.read_sql(SQL, conn)
		lines = pd.DataFrame(q)
		#loop through the lines and get all estimated costs
		for l in lines.values:
			try:
				costQuery = """
				SELECT WORKORDER_LOT_ID, EST_MATERIAL_COST AS MATERIAL, EST_LABOR_COST AS LABOR, 
				EST_BURDEN_COST AS BURDEN, EST_SERVICE_COST AS SERVICE 
				FROM REQUIREMENT WHERE WORKORDER_BASE_ID =""" + e.get() + """
				AND WORKORDER_LOT_ID = """ + str(l)[1:-1] + """GROUP BY WORKORDER_LOT_ID
				"""
			except:
				pass
		#read in all lines from quote and store them in a dataframe
		lineSQL = "SELECT LINE_NO FROM CR_QUOTE_LIN_PRICE WHERE QUOTE_ID = '" + e.get() + "'"
		q = pandas.io.sql.read_sql(lineSQL, conn)
		lines = pd.DataFrame(q)
		#loop through the lines and update percentage markups based on user input
		for l in lines:
			try:
				num = str(l)
				updatePercent = """UPDATE CR_QUOTE_LIN_PRICE SET PUR_MATL_PERCENT = """ + e2.get() + """, 
				FAB_MATL_PERCENT = """ + e3.get() + """, LABOR_PERCENT = """ + e4.get() + """, 
				BURDEN_PERCENT = """ + e5.get() + """, SERVICE_PERCENT = """ + e6.get() + """WHERE QUOTE_ID = '""" + e.get() + """' 
				AND LINE_NO = """ + num
				try:
					#commit data to database
					cursor.execute(updatePercent)
					conn.commit()
				except:
					pass
			except:
				pass
	except:
		pass

def Recalc():
	#make sure the quote number exists in the database 
	verifySQL = "SELECT ID FROM QUOTE WHERE ID IN (SELECT ID FROM QUOTE WHERE ID = '" + e.get() + "')"
	v = pandas.io.sql.read_sql(verifySQL, conn)
	D = pd.DataFrame(v)
	if D.empty:
		global top
		top = tk.Toplevel(master)
		top.title('Error')
		msg = tk.Message(top, text="Please Check Your Quote Number", width=750)
		msg.grid(row=0, column=1)
		return
	#prevent 2 from being calculated as 20 by forcing all percentages to be two digits long
	if len(e2.get()) != 2 or len(e3.get()) != 2 or len(e4.get()) != 2 or len(e5.get()) != 2 or len(e6.get()) != 2:
		top = tk.Toplevel(master)
		top.title('Error')
		msg = tk.Message(top, text="All Percentages Must Be 2 Digits Long", width=750)
		msg.grid(row=0, column=1)
		return
	#get line number and markups for quote
	try:
		top = tk.Toplevel(master)
		top.title('Recalculating')
		msg = tk.Message(top, text="Successfully Recalculated         ", width=750)
		msg.grid(row=0, column=1)
		lineSQL = "SELECT LINE_NO FROM CR_QUOTE_LIN_PRICE WHERE QUOTE_ID = '" + e.get() + "'"
		matSQL = "SELECT PUR_MATL_PERCENT FROM CR_QUOTE_LIN_PRICE WHERE QUOTE_ID = '" + e.get() + "'"
		fabSQL = "SELECT FAB_MATL_PERCENT FROM CR_QUOTE_LIN_PRICE WHERE QUOTE_ID = '" + e.get() + "'"
		labSQL = "SELECT LABOR_PERCENT FROM CR_QUOTE_LIN_PRICE WHERE QUOTE_ID = '" + e.get() + "'"
		burSQL = "SELECT BURDEN_PERCENT FROM CR_QUOTE_LIN_PRICE WHERE QUOTE_ID = '" + e.get() + "'"
		serSQL = "SELECT SERVICE_PERCENT FROM CR_QUOTE_LIN_PRICE WHERE QUOTE_ID = '" + e.get() + "'"
		#read the queries into a dataframe
		matQ = pandas.io.sql.read_sql(matSQL, conn).head(1).values
		fabQ = pandas.io.sql.read_sql(fabSQL, conn).head(1).values
		labQ = pandas.io.sql.read_sql(labSQL, conn).head(1).values
		burQ = pandas.io.sql.read_sql(burSQL, conn).head(1).values
		serQ = pandas.io.sql.read_sql(serSQL, conn).head(1).values
		#remove ugly formatting
		matQ = str(matQ).strip('[.]')
		fabQ = str(fabQ).strip('[.]')
		labQ = str(labQ).strip('[.]')
		burQ = str(burQ).strip('[.]')
		serQ = str(serQ).strip('[.]')
		#put line data into dataframe
		q = pandas.io.sql.read_sql(lineSQL, conn)
		lines = pd.DataFrame(q).values
		#loop through lines and mark up all rates based on inputs
		for l in lines:
			try:
				num = str(l)[1:-1]
				matMarkUpQuery = "SELECT 1." + e2.get() + "*EST_MATERIAL_COST FROM WORK_ORDER WHERE BASE_ID = '" + e.get() + "' AND LOT_ID = '" + num + "'"
				EPMC = str(pandas.io.sql.read_sql(matMarkUpQuery, conn).values).strip('[.]')
				updateMaterial = float(EPMC)
				roundMaterial = round(updateMaterial)

				#couldn't find a use for this since we only use purchased materials in our quotes

				# fabMarkUpQuery = "SELECT SUM(EST_MATERIAL_COST) FROM WORK_ORDER WHERE BASE_ID LIKE '%" + e.get() + "%' AND LOT_ID = '" + num + "' GROUP BY LOT_ID"
				# EFMC = str(pandas.io.sql.read_sql(fabMarkUpQuery, conn).head(1).values).strip('[.]')
				# updateFabricated = float('1.' + e.get()) * float(EPMC)
				# roundFabricated = round(updateFabricated)

				labMarkUpQuery = "SELECT 1." + e4.get() + "*EST_LABOR_COST FROM WORK_ORDER WHERE BASE_ID = '" + e.get() + "' AND LOT_ID = '" + num + "'"
				ELC = str(pandas.io.sql.read_sql(labMarkUpQuery, conn).head(1).values).strip('[.]')
				updateLabor = float(ELC)
				roundLabor = round(updateLabor)

				burMarkUpQuery = "SELECT 1." + e5.get() + "*EST_BURDEN_COST FROM WORK_ORDER WHERE BASE_ID = '" + e.get() + "' AND LOT_ID = '" + num + "'"
				EBC = str(pandas.io.sql.read_sql(burMarkUpQuery, conn).head(1).values).strip('[.]')
				updateBurden = float(EBC)
				roundBurden = round(updateBurden)

				serMarkUpQuery = "SELECT 1." + e6.get() + "*EST_SERVICE_COST FROM WORK_ORDER WHERE BASE_ID = '" + e.get() + "' AND LOT_ID = '" + num + "'"
				ESC = str(pandas.io.sql.read_sql(serMarkUpQuery, conn).head(1).values).strip('[.]')
				updateService = float(ESC)
				roundService = round(updateService)

				#calculate markups by adding all costs together, then rounding for neatness
				finalCalc = updateMaterial + updateLabor + updateBurden + updateService
				roundFinalCalc = round(finalCalc)
				strFinalCalc = str(finalCalc)
				strRoundFinalCalc = str(roundFinalCalc)

				#update in quote by replacing costs with those calculated above
				updateCosts = """UPDATE CR_QUOTE_LIN_PRICE SET CALC_UNIT_PRICE = """ + strFinalCalc + """, UNIT_PRICE = """ + strRoundFinalCalc + """WHERE QUOTE_ID = '""" + e.get() + """' 
		 							AND LINE_NO = """ + num
				cursor.execute(updateCosts)
				conn.commit()
			except:
				pass
	except:
		pass
 
def helpDoc(event=None):
	global top
	top = tk.Toplevel(master)
	top.title('Help')
	msg = tk.Message(
					top, 
					text="This program updates the rates and recalculates all lines for a quote in Visual.\
					\n\nTo use:\n\nEnter the exact quote number you wish to update.\n\nEnter percentages as two digit numbers in each box,\
					\nin the same order you would in the quote in Visual.\n\nYou should always update before recalculating.\
					\nAfter you enter the correct quote number and the rates you wish to use,\nclick the update button on the menu bar.\
					\nIf the quote number or any rates are incorrect, you will get an error message stating that.\
					\n\nAfter you have successfully updated the rates, do not clear the screen.\nClick the recalculate button.\
					\nThis program will recalculate based on the rates you have entered here, not based on the rates in Visual,\
					so it is important that the rates used to update are the same as the rates used to recalculate.\
					\n\nOnce you have updated and recalculated, you are all done. There is no need to save anything.\
					\nGo into Visual and refresh the quote.\nThe rates and prices will be updated and rounded to the nearest whole dollar.\
					\n\nShortcut keys:\n\nCtrl+R: Recalculate\nCtrl+U: Update\nF2: Reset\nEsc: Close Message\nCtrl+H: Help\nAlt+Q: Quit\
					\n\nBuilt by Michael Connell using the python programming language \nSouce code and more documentation can be found at https://github.com/MiConnell/Update-Rates\n©" + year,
					width=750
							)
	msg.pack()

#status popups
def updateStatus():
	global top
	top = tk.Toplevel(master)
	top.title('Running')
	msg = tk.Message(top, text="Updating...                                     ", width=750)
	msg.grid(row=0, column=1)
	master.after(4000, top.destroy)

def recalcStatus():
    global top
    top = tk.Toplevel(master)
    top.title('Running')
    msg = tk.Message(top, text="Recalculating...                                ", width=750)
    msg.grid(row=0, column=1)
    master.after(4000, top.destroy)

#functions to combine multiple popups
def combinedUpdate(event=None):
	updateStatus()
	master.after(3000, updateRates)

def combinedRecalc(event=None):
	recalcStatus()
	master.after(3000, Recalc)

 #closes popups
def close(event=None):
    top.destroy()

#resets screen 
def resetFromMenu(event=None):
   e.delete(0, 'end')
   e2.delete(0, 'end')
   e3.delete(0, 'end')
   e4.delete(0, 'end')
   e5.delete(0, 'end')
   e6.delete(0, 'end')
   e.focus_set()

def logIn(event=None):
	usernames = []
	passwords = []
	if en2.get() in usernames and en3.get() in passwords:
		main()
		root.destroy()
	else:
		global top
		top = tk.Toplevel(root)
		top.title('Error')
		msg = tk.Message(top, text="Incorrect Username or Password", width=750)
		msg.grid(row=0, column=1)
		return

def restart():
	master.destroy()
	security()

def security():
	global root
	root = tk.Tk()
	background = 'azure2'
	font = "Helvetica 10 bold underline"
	global en2
	en2 = tk.Entry(root)
	en2.grid(row=1, column=1, sticky="NSEW", padx=2, pady=2)
	en2.focus_set()
	global en3
	en3 = tk.Entry(root)
	en3.grid(row=2, column=1, sticky="NSEW", padx=2, pady=2)
	en3.config(show="*")
	tk.Label(root, text="User", background=background, font = font).grid(row=1)
	tk.Label(root, text="Password", background=background, font = font).grid(row=2)
	tk.Button(root, text='Log in', command=logIn).grid(row=5, column=1)
	root.bind('<Return>', logIn)
	root.bind('<Escape>', close)
	root.title("Log In")
	root.mainloop()

if __name__ == "__main__":
	security()
