# created on 22 Jan 2022
# version 2 created 16 Feb 2022, fixing decimal digit and fulfill 13 digit tax id 

from tkinter import *
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import re

root = Tk()
root.geometry('400x400')
root.title('My Receipt')

class Main_frame:
	def __init__(self, master):
		self.frame = Frame(master)
		self.menu_bar = Menu(master)

		self.info_menu = Menu(self.menu_bar, tearoff=0)
		self.info_menu.add_command(label='About Program', command=self.about)		
		self.info_menu.add_command(label='Version 2', command=self.version)
		self.info_menu.add_separator()
		self.info_menu.add_command(label='Exit', command=lambda: master.quit())
		self.menu_bar.add_cascade(label='Information', menu=self.info_menu)

	def about(self):
		root = Tk()
		root.geometry('600x200')
		root.title('About Program')
		root.config(bg='black')

		header_text = 'Functioning of My Receipt Program'
		text_1 = '1.Convert EXCEL (Receipt Inquiry Etax in OPUS) to CSV'
		text_2 = '2.Separate receipt PTE & TH'
		text_3 = '3.Checking data structure for uploading to Netbay Invchain'
		
		header_label = Label(root, text=header_text, bg='black', fg='white', font=(8))
		text_1_label = Label(root, text=text_1, bg='black', fg='white', font=(4))
		text_2_label = Label(root, text=text_2, bg='black', fg='white', font=(4))
		text_3_label = Label(root, text=text_3, bg='black', fg='white', font=(4))
		
		header_label.pack(pady=10)
		text_1_label.pack(pady=5, anchor=W)
		text_2_label.pack(pady=5, anchor=W)
		text_3_label.pack(pady=5, anchor=W)

	def version(self):
		root = Tk()
		root.geometry('600x200')
		root.title('Version 2')
		root.config(bg='black')

		text = 'Version 2\nFixing decimal digit and 13 digit tax id for both seller and buyer'

		version_label = Label(root, text=text, bg='black', fg='white', font=2)
		version_label.pack(pady=10)


class App(Main_frame):
	def __init__(self, master):
		super().__init__(master)
		
		self.btn_browse = Button(master, text='Browse', command=self.browse)
		self.btn_pte = Button(master, text='Convert PTE', state=DISABLED, command= lambda: self.convert(self.df_pte, 'PTE'))
		self.btn_th = Button(master, text='Convert TH', state=DISABLED, command= lambda: self.convert(self.df_th, 'TH'))
		self.is_office_columnB = False

	def browse(self):
		open_file = filedialog.askopenfilename()

		try:
			self.df = pd.read_excel(open_file)
			
			self.df['Prefix Receipt'] = self.df['docId'].apply(lambda x: separate_prefix(x))		
		except ValueError as e:
			print(f'{e} File format is not excel file')
			messagebox.showerror('Error', 'Not the correct file')
		except FileNotFoundError as e:
			print(f'{e} Did not select file')
		except KeyError as e:
			print(f'{e} Not Receipt Inquiry Etax file')
			messagebox.showerror('Error', 'Not the correct file')
		except Exception as e:
			print(e)
		else:
			filt_TH = (self.df['Prefix Receipt'] == 'BKKL') | (self.df['Prefix Receipt'] == 'EPBKL') | (self.df['Prefix Receipt'] == 'LCBL') | (self.df['Prefix Receipt'] == 'SGZL')

			# make DataFrame for PTE
			self.df_pte = self.df.drop(index = self.df[filt_TH].index)
			self.df_pte['taxIDBR'] = self.df_pte['taxIDBR'].map('{:0>13}'.format)
			self.df_pte['taxIDSL'] = self.df_pte['taxIDSL'].map('{:0>13}'.format)
			self.df_pte['Receipt\nApplyAMT'] = self.df_pte['Receipt\nApplyAMT'].map('{:,.2f}'.format)
			self.df_pte['chargeAmount'] = self.df_pte['chargeAmount'].map('{:,.2f}'.format)
			self.df_pte['netLine\nTotalAmount'] = self.df_pte['netLine\nTotalAmount'].map('{:,.2f}'.format)
			self.df_pte['line\nTotalAmount'] = self.df_pte['line\nTotalAmount'].map('{:,.2f}'.format)
			self.df_pte['taxBasis\nTotalAmount'] = self.df_pte['taxBasis\nTotalAmount'].map('{:,.2f}'.format)
			self.df_pte['calculate\nRate'] = self.df_pte['calculate\nRate'].map('{:,.2f}'.format)
			self.df_pte['tax\nTotalAmount'] = self.df_pte['tax\nTotalAmount'].map('{:,.2f}'.format)
			self.df_pte['grand\nTotalAmount'] = self.df_pte['grand\nTotalAmount'].map('{:,.2f}'.format)			
			self.df_pte.replace(['nan'], np.nan, inplace=True)			
			try:		
				self.df_pte.drop(columns=['Unnamed: 0'], inplace=True) 
				self.is_office_columnB = True
			except KeyError as e:
				print(e)
				messagebox.showerror('Error', 'Office must be in column B for PTE')
				self.is_office_columnB = False
			
			# make DataFrame for TH
			self.df_th = self.df.drop(index = self.df[~filt_TH].index)
			self.df_th['taxIDBR'] = self.df_th['taxIDBR'].map('{:0>13}'.format)
			self.df_th['taxIDSL'] = self.df_th['taxIDSL'].map('{:0>13}'.format)
			self.df_th['calculate\nRate'] = self.df_th['calculate\nRate'].map('{:,.2f}'.format)
			self.df_th['grand\nTotalAmount'] = self.df_th['grand\nTotalAmount'].map('{:,.2f}'.format)
			self.df_th.replace(['nan'], np.nan, inplace=True)
			try:
				self.df_th.drop(columns=['Unnamed: 0'], inplace=True)
				self.is_office_columnB = True
			except KeyError as e:
				print(e)
				messagebox.showerror('Error', 'Office must be in column B for TH')
				self.is_office_columnB = False
			
			# checking data structure
			if self.is_office_columnB == True:
				messagebox.showinfo('Message', 'Ready to Convert :-)') 
				if 'ok':
					self.btn_pte['state'] = 'normal'
					self.btn_th['state'] = 'normal'
					self.btn_browse['state'] = 'disabled'
			elif self.is_office_columnB == False:
				messagebox.showwarning('Warning', 'No file to convert')
			
	def convert(self, df, name):		
		f_ext = [('csv', '*.csv')]
		
		try:
			convert_path = filedialog.asksaveasfilename(filetypes=f_ext, defaultextension=f_ext, initialfile= f'Receipt Inquiry {name}')			
			df.to_csv(convert_path, index=False)
		except FileNotFoundError as e:
			print(e)		
		else:
			messagebox.showinfo('Message', f'Convert {name} Successful')
			
	def draw(self):
		self.frame.pack(fill='both', expand=True)				
		self.btn_browse.place(x=130, y=50, width=150, height=50)
		self.btn_pte.place(x=130, y=120, width=150, height=50)
		self.btn_th.place(x=130, y=190, width=150, height=50)


class Developer(Main_frame):
	def __init__(self, master, name):
		super().__init__(master)
		self.label = Label(master, text=name, fg='grey')

	def draw(self):
		self.label.pack(anchor=SE, padx=10, pady=10)
		
		
def separate_prefix(text):
	result = re.findall(r'BKKO|BKKL|BKKI|EPBKO|EPBKI|EPBKL|LCBO|LCBI|LCBL|SGZO|SGZI|SGZL\.?', text)
	return ' '.join(result)

def main():	
	app.draw()	
	developer.draw()

	root.mainloop()

main_frame = Main_frame(root)
app = App(root)
developer = Developer(root, 'V.2 Created by: Natcha Phonkamhaeng')
root.config(menu=main_frame.menu_bar)

if __name__ == '__main__':
	main()




	








