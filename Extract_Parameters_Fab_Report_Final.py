import glob
import os
import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog, StringVar, Label, Button, Entry
from pathlib import Path
import shutil
import xlwings as xw

global choice
global varList
global Template

dir_path = os.path.dirname(os.path.realpath(__file__))
# # print(dir_path)

Fab_Report_Template = str(dir_path) + "\\" + "Fab_Report_Template.xlsx"

# Template = pd.read_excel(Fab_Report_Template, index_col=False)
# # print(Template)

def display_selected_params(choice):
	choice = varList.get()
	# # print(choice)

def display_selected_slot(choice):
	choice = varSlot.get()
	# # print(choice)

# --- Open Thickness File---

def Open_Thk_File():
	# Thk_Template = copy.deepcopy(Template)
	# # print("Thk_Template:")
	# # print(Thk_Template)

	Thk_dataframe = pd.read_excel(
		Thk_Open_FileName, skiprows=4, index_col=False)
	Thk_dataframe_sorted = Thk_dataframe.fillna(method='ffill').iloc[:, 2:]
	# # print(Thk_dataframe_sorted)

	Recipe_List = Thk_dataframe_sorted["Jobfile Name"].dropna().unique()
	# print(Recipe_List)

	shutil.copyfile(Fab_Report_Template, Thk_Save_FullFileName)

	workbook = xw.Book(Thk_Save_FullFileName)
	worksheet = workbook.sheets[0]
	worksheet.name = varList.get()
	worksheet["K1"].value = varSlot.get()

	Thk_Compiled_Df = pd.DataFrame()

	for i in Recipe_List:
		# # print(i)
		Thk_Recipe = Thk_dataframe_sorted[Thk_dataframe_sorted["Jobfile Name"] == i]
		# # print(Thk_Recipe)

		Thk_Recipe_Slot = Thk_Recipe[Thk_Recipe["Slot.No"] == varSlot.get()]
		# # print(Thk_Recipe_Slot)

		if Thk_Recipe_Slot.empty:
			continue

		Thk_Recipe_Slot_df = Thk_Recipe_Slot.iloc[:, 3:]
		# # print("This is slot values")
		# # print(Thk_Recipe_Slot_df)

		Thk_Recipe_Slot_df_Transpose = Thk_Recipe_Slot_df.reset_index(
			drop=True).transpose().reset_index(drop=True).round(decimals=2)
		# # print(Thk_Recipe_Slot_df_Transpose)

		Thk_Transpose_Row = int(Thk_Recipe_Slot_df_Transpose.shape[0])
		# # print(Thk_Transpose_Row)
		Thk_Transpose_Col = int(Thk_Recipe_Slot_df_Transpose.shape[1])
		# # print(Thk_Transpose_Col)

		Site_List = []
		Site_Count = 1
		for j in range(Thk_Transpose_Row):
			Site_List.append(Site_Count)
			Site_Count += 1
		# # print(Site_List)
		Site_List_Df = pd.DataFrame(Site_List).reset_index(drop=True)
		# print(Site_List_Df)

		Recipe_Jobfile_Row_List = [i]*Thk_Transpose_Row
		Recipe_Jobfile_Row_List_Df = pd.DataFrame(
			Recipe_Jobfile_Row_List).reset_index(drop=True)
		# # print(Recipe_Jobfile_Row_List_Df)

		Thk_Combined_Jobfile_And_Values = pd.concat(
			[Site_List_Df, Recipe_Jobfile_Row_List_Df, Thk_Recipe_Slot_df_Transpose], axis=1)

		Thk_Compiled_Df = pd.concat(
			[Thk_Compiled_Df, Thk_Combined_Jobfile_And_Values], axis=0).dropna()

	# print(Thk_Compiled_Df)

	Thk_Compiled_Df_1 = Thk_Compiled_Df.iloc[:, 0]
	Thk_Compiled_Df_2 = Thk_Compiled_Df.iloc[:, 1]
	Thk_Compiled_Df_3 = Thk_Compiled_Df.iloc[:, 2]

	worksheet["B2"].options(
		index=False, header=False).value = Thk_Compiled_Df_1
	worksheet["C2"].options(
		index=False, header=False).value = Thk_Compiled_Df_2
	worksheet["F2"].options(
		index=False, header=False).value = Thk_Compiled_Df_3

# --- For Open CD File ---

def Open_CD_File():
	global CD_Folder_List_Names
	# print(CD_Folders)

	CD_all_files_combined_dataframe = pd.DataFrame()

	shutil.copyfile(Fab_Report_Template, CD_Save_FullFileName)

	CD_workbook = xw.Book(CD_Save_FullFileName)
	CD_worksheet = CD_workbook.sheets[0]
	CD_worksheet.name = varList.get()
	CD_worksheet["K1"].value = varSlot.get()

	CD_Folder_Count = 0

	for file in CD_Folders:
		# print("file")
		# print(file)
		CD_dataframe = pd.read_excel(
		    file, skiprows=4, index_col=False, sheet_name='Average')
		CD_dataframe_sorted = CD_dataframe.fillna(method='ffill').iloc[:, 2:]
		# # print(CD_dataframe_sorted)

		Structure_List = CD_dataframe_sorted["Structure"].dropna().unique()
		# # print(Structure_List)

		CD_Compiled_Df = pd.DataFrame()

		for i in Structure_List:
			CD_Structure = CD_dataframe_sorted[CD_dataframe_sorted["Structure"] == i]
			# # print(CD_Structure)

			CD_Slot = "Slot " + str(varSlot.get())
			CD_Structure_Slot = CD_Structure[CD_Structure["Slot"] == CD_Slot]
			# print(CD_Structure_Slot)

			if CD_Structure_Slot.empty:
				continue

			CD_Structure_Slot_df = CD_Structure_Slot.iloc[:, 2:7]
			# print("Structure Slots: ")
			# print(CD_Structure_Slot_df)

			CD_Structure_Slot_df_Transpose = CD_Structure_Slot_df.reset_index(
			    drop=True).transpose().reset_index(drop=True).round(decimals=2)
			# print(CD_Structure_Slot_df_Transpose)

			CD_Transpose_Row = int(CD_Structure_Slot_df_Transpose.shape[0])
			CD_Transpose_Col = int(CD_Structure_Slot_df_Transpose.shape[1])

			Site_List = []
			Site_Count = 1
			for j in range(CD_Transpose_Row):
				Site_List.append(Site_Count)
				Site_Count += 1
			# # print(Site_List)
			Site_List_Df = pd.DataFrame(Site_List).reset_index(drop=True)
			# print(Site_List_Df)

			Structure_Row_List = [i]*CD_Transpose_Row
			Structure_Row_List_Df = pd.DataFrame(
			    Structure_Row_List).reset_index(drop=True)

			CD_Combined_Structure_And_Values = pd.concat(
			    [Site_List_Df, Structure_Row_List_Df, CD_Structure_Slot_df_Transpose], axis=1)

			CD_Compiled_Df = pd.concat(
			    [CD_Compiled_Df, CD_Combined_Structure_And_Values], axis=0).dropna().reset_index(drop=True)

		# print(CD_Compiled_Df)
		# print("---")
		
		CD_Folder_Name = CD_Folder_List_Names[CD_Folder_Count]
		# print("Hi")
		# print(CD_Folder_Name)
		CD_Folder_Row = [CD_Folder_Name]*(CD_Compiled_Df.shape[0])
		CD_Folder_Row_Df = pd.DataFrame(CD_Folder_Row)
		# print("hihi")
		# print(CD_Folder_Row_Df)
		# print("hihihi")
		CD_Final_Compiled_Df = pd.concat([CD_Folder_Row_Df, CD_Compiled_Df], axis=1)
		# print(CD_Final_Compiled_Df)
		CD_Folder_Count += 1

		CD_all_files_combined_dataframe = pd.concat([CD_all_files_combined_dataframe, CD_Final_Compiled_Df], axis=0)



	# print("final final")
	# print(CD_all_files_combined_dataframe)
	CD_Final_Compiled_Df_1 = CD_all_files_combined_dataframe.iloc[:, 0]
	CD_Final_Compiled_Df_2 = CD_all_files_combined_dataframe.iloc[:, 1]
	CD_Final_Compiled_Df_3 = CD_all_files_combined_dataframe.iloc[:, 2]
	CD_Final_Compiled_Df_4 = CD_all_files_combined_dataframe.iloc[:, 3]

	CD_worksheet["A2"].options(
		index=False, header=False).value = CD_Final_Compiled_Df_1
	CD_worksheet["B2"].options(
		index=False, header=False).value = CD_Final_Compiled_Df_2
	CD_worksheet["C2"].options(
		index=False, header=False).value = CD_Final_Compiled_Df_3
	CD_worksheet["F2"].options(
		index=False, header=False).value = CD_Final_Compiled_Df_4


# --- ET ---
def Open_ET_File():
	ET_dataframe = pd.read_excel(
	    ET_File, skiprows=1, index_col=False, sheet_name='Summary')
	# # print(ET_dataframe)

	ET_dataframe_sorted = ET_dataframe.fillna(method='ffill').iloc[:,2:]
	# # print("ET DF")
	# # print(ET_dataframe_sorted)

	Device_List = ET_dataframe_sorted["Device"].dropna().unique()
	# print(Device_List)

	shutil.copyfile(Fab_Report_Template,ET_Save_FullFileName)

	ET_workbook = xw.Book(ET_Save_FullFileName)
	ET_worksheet = ET_workbook.sheets[0]
	ET_worksheet.name = varList.get()
	ET_worksheet["K1"].value = varSlot.get()

	ET_Compiled_Df = pd.DataFrame()

	ET_dataframe_sorted_device = ET_dataframe_sorted[ET_dataframe_sorted["Device"].isin(['2P','XH3'])].reset_index(drop=True)
	# # print(ET_dataframe_sorted_device)

	ET_dataframe_sorted_device_slot = ET_dataframe_sorted_device[ET_dataframe_sorted_device["Slot"] == varSlot.get()].reset_index(drop=True)
	# # print(ET_dataframe_sorted_device_slot)

	ET_dataframe_sorted_device_slot_2P = (ET_dataframe_sorted_device_slot[ET_dataframe_sorted_device_slot["Device"].isin(['2P'])].iloc[:,2:7]*1000000000).round(decimals=2)

	# # print(ET_dataframe_sorted_device_slot_2P)
	ET_dataframe_sorted_device_slot_2P_T = ET_dataframe_sorted_device_slot_2P.reset_index(drop=True).transpose().reset_index(drop=True).round(decimals=2)
	ET_2P_Row = ["2P"]*ET_dataframe_sorted_device_slot_2P_T.shape[0]
	ET_2P_Row_Df = pd.DataFrame(ET_2P_Row).reset_index(drop=True)
	# print(ET_2P_Row_Df)

	Site_List_2P = []
	Site_Count_2P = 1
	for j in range(ET_2P_Row_Df.shape[0]):
		Site_List_2P.append(Site_Count_2P)
		Site_Count_2P += 1
	Site_List_Df_2P = pd.DataFrame(Site_List_2P).reset_index(drop=True)

	ET_2P_Site_Combined = pd.concat([Site_List_Df_2P, ET_2P_Row_Df], axis=1).dropna().reset_index(drop=True)

	ET_dataframe_sorted_device_slot_XH3 = (ET_dataframe_sorted_device_slot[ET_dataframe_sorted_device_slot["Device"].isin(['XH3'])].iloc[:,2:7])
	ET_dataframe_sorted_device_slot_XH3_T = ET_dataframe_sorted_device_slot_XH3.reset_index(drop=True).transpose().reset_index(drop=True).round(decimals=2)
	ET_XH3_Row = ["XH3"]*ET_dataframe_sorted_device_slot_XH3_T.shape[0]
	ET_XH3_Row_Df = pd.DataFrame(ET_XH3_Row).reset_index(drop=True)

	Site_List_XH3 = []
	Site_Count_XH3 = 1
	for j in range(ET_XH3_Row_Df.shape[0]):
		Site_List_XH3.append(Site_Count_XH3)
		Site_Count_XH3 += 1
	Site_List_Df_XH3 = pd.DataFrame(Site_List_XH3).reset_index(drop=True)

	ET_XH3_Site_Combined = pd.concat([Site_List_Df_XH3, ET_XH3_Row_Df], axis=1).dropna().reset_index(drop=True)

	ET_Device_Site_Combined_Final = pd.concat([ET_2P_Site_Combined, ET_XH3_Site_Combined], axis=0).dropna().reset_index(drop=True)
	# # print(ET_Device_Site_Combined_Final)

	ET_dataframe_sorted_device_slot_T_combined = pd.concat([ET_dataframe_sorted_device_slot_2P_T,ET_dataframe_sorted_device_slot_XH3_T], axis=0).dropna().reset_index(drop=True)
	# # print(ET_dataframe_sorted_device_slot_T_combined)

	ET_worksheet["B2"].options(index=False, header=False).value = ET_Device_Site_Combined_Final
	ET_worksheet["F2"].options(index=False, header=False).value = ET_dataframe_sorted_device_slot_T_combined

# --- For All Parameters Popup ---

def popup_window():
	# Thickness Popup
	if varList.get() == Parameters[0]:
		window = Toplevel()
		label = Label(
			window, text="Extract Thickness Data into Fab Report Format!")
		label.pack()

		Label_Space = Label(window, text="              ")
		Label_Space.pack()

		def Thk_Open_File():
			global Thk_Open_FileName
			thk_open_file_entry.config(state=NORMAL)
			thk_open_file_entry.delete(0, 'end')
			# ask for user input, return file names in a tuple
			Thk_Open_FileName = filedialog.askopenfilename(
				filetypes=[('Excel', '*.xlsx'), ('CSV Files', '*.csv')])
			thk_open_file_entry.insert(tk.END, Thk_Open_FileName)
			Thk_basename = Path(Thk_Open_FileName).stem
			# print(Thk_basename)
			window.lift()

		Thk_Open_Button = Button(
			window, text="Open Compiled Thk File", command=Thk_Open_File)
		Thk_Open_Button.pack()

		Thk_Open_Entry = StringVar(value='Open Compiled Thk Data Location')
		thk_open_file_entry = Entry(
			window, width=50, textvariable=Thk_Open_Entry, foreground='gray', justify=CENTER)
		thk_open_file_entry.config(state=DISABLED)
		thk_open_file_entry.pack()

		def Thk_Save_File():
			global Thk_Save_FileName
			global Thk_Save_FullFileName
			save_file_entry.config(state=NORMAL)
			save_file_entry.delete(0, 'end')
			Thk_Save_FileName = filedialog.asksaveasfilename(
				filetypes=[('Excel', '*.xlsx')], initialdir=os.path.expanduser("~/Desktop"))
			save_file_entry.insert(tk.END, Thk_Save_FileName)
			Thk_Save_FullFileName = Thk_Save_FileName+".xlsx"
			window.lift()

		Label_Space = Label(window, text="              ")
		Label_Space.pack()

		Thk_Save_Button = Button(
			window, text="Save File", command=Thk_Save_File)
		Thk_Save_Button.pack()

		Thk_Save_Entry = StringVar(value='Saved Thk Data Location')
		save_file_entry = Entry(
			window, width=50, textvariable=Thk_Save_Entry, foreground='gray', justify=CENTER)
		save_file_entry.config(state=DISABLED)
		save_file_entry.pack()

		Label_Space = Label(window, text="              ")
		Label_Space.pack()

		button_select = Button(
			window, text="Press this to start!", command=Open_Thk_File)
		button_select.pack()

		Label_Space = Label(window, text="              ")
		Label_Space.pack()

		button_close = Button(window, text="Close", command=window.destroy)
		button_close.pack()

	# CD Popup
	elif varList.get() == Parameters[1]:
		window = Toplevel()
		label = Label(window, text="CD Window!")
		label.pack()

		Label_Space = Label(window, text="              ")
		Label_Space.pack()

		def CD_Open_File():
			global CD_Folders
			global CD_basename
			global CD_Folder_List_Names
			CD_Folders = filedialog.askopenfilenames(
				filetypes=[('Excel', '*.xlsx'), ('CSV Files', '*.csv')])
			# filetypes = [('Excel', '*.xlsx'),('CSV Files','*.csv')]
			CD_Folder_List_Names = []
			# print(CD_Folders)

			CD_open_file_entry.config(state=NORMAL)
			CD_open_file_entry.delete(0, 'end')

			for file in CD_Folders:
				CD_basename = Path(file).stem
				# # print(CD_basename)
				CD_Folder_List_Names.append(CD_basename)
			# # print(CD_Folder_List_Names)

			CD_open_file_entry.insert(tk.END, CD_Folder_List_Names)
			window.lift()

		button_folder = Button(
			window, text="Select All CD Excel Files", command=CD_Open_File)
		button_folder.pack()

		CD_Open_Entry_Text = StringVar(value='Open Compiled CD Data Location')
		CD_open_file_entry = Entry(
			window, width=50, textvariable=CD_Open_Entry_Text, foreground='gray', justify=CENTER)
		CD_open_file_entry.config(state=DISABLED)
		CD_open_file_entry.pack()

		Label_Space = Label(window, text="              ")
		Label_Space.pack()

		def CD_Save_File():
			global CD_Save_FileName
			global CD_Save_FullFileName
			save_file_entry.config(state=NORMAL)
			save_file_entry.delete(0, 'end')
			CD_Save_FileName = filedialog.asksaveasfilename(
				filetypes=[('Excel', '*.xlsx')], initialdir=os.path.expanduser("~/Desktop"))
			save_file_entry.insert(tk.END, CD_Save_FileName)
			CD_Save_FullFileName = CD_Save_FileName+".xlsx"
			window.lift()

		CD_Save_Button = Button(window, text="Save File", command=CD_Save_File)
		CD_Save_Button.pack()

		CD_Save_Entry = StringVar(value='Saved CD Data Location')
		save_file_entry = Entry(
			window, width=50, textvariable=CD_Save_Entry, foreground='gray', justify=CENTER)
		save_file_entry.config(state=DISABLED)
		save_file_entry.pack()

		Label_Space = Label(window, text="              ")
		Label_Space.pack()

		button_select = Button(
			window, text="Press this to start!", command=Open_CD_File)
		button_select.pack()

		Label_Space = Label(window, text="              ")
		Label_Space.pack()

		button_close = Button(window, text="Close", command=window.destroy)
		button_close.pack()

	# OT Popup (No need for OT, just use Sola's macro)
	# elif varList.get() == Parameters[2]:
	# 	window = Toplevel()
	# 	label = Label(window, text="OT Window!")
	# 	label.pack()
	# 	Label_Space = Label(window, text="              ")
	# 	Label_Space.pack()

	# ET Popup
	elif varList.get() == Parameters[2]:
		window = Toplevel()
		label = Label(window, text="ET Window!")
		label.pack()

		Label_Space = Label(window, text="              ")
		Label_Space.pack()

		def ET_Open_File():
			global ET_File
			global ET_basename
			ET_File = filedialog.askopenfilename(
				filetypes=[('Excel', '*.xlsx'), ('CSV Files', '*.csv')])
			# print(ET_File)

			ET_open_file_entry.config(state=NORMAL)
			ET_open_file_entry.delete(0, 'end')
			ET_basename = Path(ET_File).stem
			ET_open_file_entry.insert(tk.END, ET_basename)
			window.lift()

		button_folder = Button(
			window, text="Select All ET Excel Files", command=ET_Open_File)
		button_folder.pack()

		ET_Open_Entry_Text = StringVar(value='Open Compiled ET Data Location')
		ET_open_file_entry = Entry(
			window, width=50, textvariable=ET_Open_Entry_Text, foreground='gray', justify=CENTER)
		ET_open_file_entry.config(state=DISABLED)
		ET_open_file_entry.pack()

		Label_Space = Label(window, text="              ")
		Label_Space.pack()

		def ET_Save_File():
			global ET_Save_FileName
			global ET_Save_FullFileName
			save_file_entry.config(state=NORMAL)
			save_file_entry.delete(0, 'end')
			ET_Save_FileName = filedialog.asksaveasfilename(
				filetypes=[('Excel', '*.xlsx')], initialdir=os.path.expanduser("~/Desktop"))
			save_file_entry.insert(tk.END, ET_Save_FileName)
			ET_Save_FullFileName = ET_Save_FileName+".xlsx"
			window.lift()

		ET_Save_Button = Button(window, text="Save File", command=ET_Save_File)
		ET_Save_Button.pack()

		ET_Save_Entry = StringVar(value='Saved ET Data Location')
		save_file_entry = Entry(
			window, width=50, textvariable=ET_Save_Entry, foreground='gray', justify=CENTER)
		save_file_entry.config(state=DISABLED)
		save_file_entry.pack()

		Label_Space = Label(window, text="              ")
		Label_Space.pack()

		button_select = Button(
			window, text="Press this to start!", command=Open_ET_File)
		button_select.pack()

		Label_Space = Label(window, text="              ")
		Label_Space.pack()

		button_close = Button(window, text="Close", command=window.destroy)
		button_close.pack()

App = tk.Tk()
App.title("Fab Report Extract")
App.geometry("460x255")

Label_Start = Label(App, text="First, select the parameter to extract:")
Label_Start.grid(row=0, column=0, sticky=W, padx=2, pady=2)

varList = StringVar(App)
varList.set("Select Me!")
Parameters = ["Thickness",
			  "Critical Dimension (CD)", "Electrical Test (ET)"]
Option_Params = OptionMenu(App, varList, *Parameters,
						   command=display_selected_params)
Option_Params.grid(row=1, column=0, sticky=W, padx=2, pady=2)

Submit_Button = Button(App, text="Confirm", command=popup_window)
Submit_Button.grid(row=1, column=1, sticky=W, padx=2, pady=2)

Label_Slot = Label(App, text="Please select one slot to work on at a time:")
Label_Slot.grid(row=2, column=0, sticky=W, padx=2, pady=2)

varSlot = IntVar(App)
Slots = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13,
		 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25]
Option_Slots = OptionMenu(App, varSlot, *Slots, command=display_selected_slot)
Option_Slots.grid(row=3, column=0, sticky=W, padx=2, pady=2)

Label_Space = Label(App, text="              ")
Label_Space.grid(row=4, column=0, sticky=W, padx=2, pady=2)

Label1 = Label(App, text="First use Sola's Macro to compile all your data. ")
Label1.grid(row=5, column=0, sticky=W, padx=2, pady=2)

Label_Thk = Label(App, text="For Thk, should select all and put in ONE compiled FILE")
Label_Thk.grid(row=6, column=0, sticky=W, padx=2, pady=2)

Label_CD = Label(App, text="For CD, compile it into EACH LAYER (If got Slab and SiN CD, then 2 files)")
Label_CD.grid(row=7, column=0, sticky=W, padx=2, pady=2)

Label_ET = Label(App, text="For ET, select all and put in ONE compiled FILE")
Label_ET.grid(row=8, column=0, sticky=W, padx=2, pady=2)

App.mainloop()

