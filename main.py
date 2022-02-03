# created on 22 Jan 2022

from tkinter import *
from tkinter import filedialog, messagebox
import pandas as pd
import re

root = Tk()
root.geometry('400x400')
root.title('My Receipt')

class Main_frame:
	def __init__(self, master):
		self.frame = Frame(master)

	def draw(self):
		self.frame.pack(fill='both', expand=True)


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
			try:		
				self.df_pte.drop(columns=['Unnamed: 0'], inplace=True) 
				self.is_office_columnB = True
			except KeyError as e:
				print(e)
				messagebox.showerror('Error', 'Office must be in column B for PTE')
				self.is_office_columnB = False
			
			# make DataFrame for TH
			self.df_th = self.df.drop(index = self.df[~filt_TH].index)
			try:
				self.df_th.drop(columns=['Unnamed: 0'], inplace=True)
				self.is_office_columnB = True
			except KeyError as e:
				print(e)
				messagebox.showerror('Error', 'Office must be in column B for TH')
				self.is_office_columnB = False
			
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
	main_frame = Main_frame(root)
	main_frame.draw()

	app = App(root)
	app.draw()

	developer = Developer(root, 'V.1 Created by: Natcha Phonkamhaeng')
	developer.draw()

	root.mainloop()

if __name__ == '__main__':
	main()




	








