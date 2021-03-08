import tkinter as tk
from tkinter import *
from calculator import Excel_calculator

root = Tk()

#######################
version = "v" + "0.8"
#######################
class color:
	dark = "#31bc63"
	light = "#a3ffc3"
	primary = "#6def92"
#######################

class option:
	calcSaturated = tk.IntVar()
	calcMono = tk.IntVar()
	calcTrans = tk.IntVar()
	calcPoly = tk.IntVar()
	allSelected = tk.IntVar()

	def update(self):
		self.calcSaturated.set(self.allSelected.get())
		self.calcMono.set(self.allSelected.get())
		self.calcTrans.set(self.allSelected.get())
		self.calcPoly.set(self.allSelected.get())
		self.allSelected.set(self.allSelected.get())

options = option()

# this is a function to get the user input from the text input box
def getInputBoxValue():
	userInput = filenameInput.get()
	return userInput




#####################################################################
# Get the info and move on to the other file to do the calculations #
#####################################################################
def calculate():
	try:
		output = Excel_calculator(filenameInput.get(), options)

	except Exception as e:
		output = e

	outputText.set(output)



######################
# Main tkinter setup #
######################

# this is the declaration of the variable associated with the checkbox
# calcSaturated = tk.IntVar()
# calcMono = tk.IntVar()
# calcTrans = tk.IntVar()
# calcPoly = tk.IntVar()

# This is the section of code which creates the main window
root.geometry('290x390')
root.configure(background='#FFFFFF')
root.title('Excel Calculator')

# This is the section of code which creates the a label
Label(root, text='Enter the name of the Input file', bg='#FFFFFF', font=('arial', 14, 'normal')).place(x=12, y=13)

# This is the section of code which creates a text input box
filenameInput = Entry(root)
filenameInput.place(x=12, y=63)

# This is the section of code which creates the a label
Label(root, text='.xlsx', bg='#FFFFFF', font=('arial', 10, 'normal')).place(x=152, y=63)

#########################################
# Checkboxes for aditional calculations #
#########################################

Checkbutton(root, text='Saturated', variable=options.calcSaturated, bg=color.dark,font=('arial', 12, 'normal')).place(x=12, y=103)

Checkbutton(root, text='Mono', variable=options.calcMono, bg=color.dark,font=('arial', 12, 'normal')).place(x=12, y=133)

Checkbutton(root, text='Poly', variable=options.calcPoly, bg=color.dark,font=('arial', 12, 'normal')).place(x=12, y=163)

Checkbutton(root, text='Trans', variable=options.calcTrans, bg=color.dark,font=('arial', 12, 'normal')).place(x=12, y=193)

Checkbutton(root, text='All', variable=options.allSelected, bg=color.light,font=('arial', 12, 'normal'), command=options.update).place(x=12, y=223)




# Proceed Button
Button(root, text='Proceed', bg=color.dark, font=('arial', 14, 'normal'), command=calculate).place(relx=0.5, rely=0.9, anchor=CENTER)


# Output text
outputText = StringVar()
outputLablel = Label(root, text="", textvariable = outputText, bg='#FFFFFF', font=('arial', 14, 'normal')).place(relx=0.5, rely=0.8, anchor=CENTER)

# Version label
Label(root, text=version, bg='#FFFFFF', font=('arial', 10, 'normal')).place(rely=0.98, anchor=W)


root.mainloop()
