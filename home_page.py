import tkinter as tk
import tkinter.ttk as ttk
from encryptor import Encryptor
from decryptor import Decryptor
from tkinter import messagebox, filedialog
import os
import random
import string
import pyAesCrypt
from openpyxl import *
import tempfile
from showdirs import Showdirs
from Gdrive import *

class Homepage:

	def __init__( self, root, master, user, userId,password, dirname, dirlists, filelists ):
		# build ui
		self.root = root
		self.mainframe = master
		self.user = user
		self.userId = userId
		self.password = password
		self.dirname = dirname
		self.dirlists = dirlists
		self.filelists = filelists
		self.homepage_F = ttk.Frame( master )
		hide_B = ttk.Button( self.homepage_F )
		hide_B.config( text='Add Folder' )
		hide_B.grid( padx='60', pady='10' )
		hide_B.configure( command=self.hide )
		show_B = ttk.Button( self.homepage_F )
		show_B.config( text='Show' )
		show_B.grid( column='1', row='0', padx='60', pady='10' )
		show_B.configure( command=self.show )
		backup_B = ttk.Button( self.homepage_F )
		backup_B.config( text='Backup' )
		backup_B.grid( column='0', row='1', padx='10', pady='10' )
		backup_B.configure( command=self.back_up )
		restore_B = ttk.Button( self.homepage_F )
		restore_B.config( text='Restore' )
		restore_B.grid( column='1', row='1', padx='10', pady='10' )
		restore_B.configure( command=self.restore )
		# changepass_B = ttk.Button( self.homepage_F )
		# changepass_B.config( text='Change\nPassword' )
		# changepass_B.grid( column='0', row='2', padx='10', pady='10' )
		# changepass_B.configure( command=self.change_pass )
		# deleteuser_B = ttk.Button( self.homepage_F )
		# deleteuser_B.config( text='Delete\nUser' )
		# deleteuser_B.grid( column='1', row='2', padx='10', pady='10' )
		# deleteuser_B.configure( command=self.delete_user )
		logout_B = ttk.Button( self.homepage_F )
		logout_B.config( text='Logout' )
		logout_B.grid( column='0', row='3', padx='10', pady='10' )
		logout_B.configure( command=self.logout )
		self.homepage_F.config( height='200', relief='flat', width='200' )
		self.homepage_F.place( anchor='nw', height='500', width='500', x='0', y='10' )

		# Main widget
		self.mainwindow = self.homepage_F
		self.config()
		hide_B.focus_set()

	def config( self ):
		self.mainframe.config( text=self.user )
		# self.home = os.path.abspath( os.path.join( os.path.expanduser( '~' ), "AppData\\Local\\Programs\\Invisible" ) )
		self.home = "Personal"
		self.dest_dir = os.path.join( self.home, self.dirname )
		self.bufferSize = 64 * 1024

	def random_str( self, n ):
		return ( ''.join( random.choices( string.ascii_uppercase + string.digits + string.ascii_lowercase, k=n ) ) )

	def hide( self ):
		self.src_dir = filedialog.askdirectory( title="Select Folder" )
		if self.src_dir != "":
			self.mainframe.config( text="Hidding..." )
			self.mainframe.update()
			self.src_dir = os.path.abspath( self.src_dir )
			Enc = Encryptor( self.src_dir, self.dest_dir, self.dirlists, self.filelists, self.password )
			Enc.encrypt()
			self.mainframe.config( text=self.user )
			self.mainframe.update()

		else:
			messagebox.showerror( 'Error', "Folder not Selected !" )
		pass

	def show( self ):
		# home = os.path.abspath( os.path.join( os.path.expanduser( '~' ), "AppData\\Local\\Programs\\Invisible" ) )
		home = "Invisible"
		sd = Showdirs(
		    self.root, self.mainframe, home, self.user, self.password, self.dirname, self.dirlists, self.filelists
		    )

	def back_up( self ):
		src_dir = filedialog.askdirectory( title="Select Source Folder" )
		dest_dir = filedialog.askdirectory( title="Select Destination Folder" )
		#print( src_dir, dest_dir )
		if src_dir == dest_dir:
			messagebox.showerror( 'Error', "Source and Destination Folders are Never be same !" )
		elif src_dir != "" and dest_dir != "":
			self.mainframe.config( text="Hidding..." )
			self.mainframe.update()
			#print( src_dir, dest_dir )
			Enc = Encryptor(
			    src_dir, os.path.join( dest_dir, self.dirname ), self.dirlists, self.filelists, self.password
			    )
			Enc.encrypt()
			self.mainframe.config( text=self.user )
			self.mainframe.update()
		else:
			messagebox.showerror( "Error", "Folder not Selected !" )

	def restore( self ):
		home = filedialog.askdirectory( title="Select Restore Folder" )
		if home == "":
			messagebox.showerror( "Error", "Folder not Selected !" )
		elif not os.path.exists( os.path.join( home, self.dirname, self.dirlists ) ):
			messagebox.showerror( "Error", "No restore point found !" )
		else:
			sd = Showdirs(
			    self.root, self.mainframe, home, self.user, self.password, self.dirname, self.dirlists, self.filelists
			    )

	# def change_pass( self ):
	# 	pass

	# def delete_user( self ):
	# 	pass

	def logout( self ):
		self.mainframe.config( text="Invisible" )
		self.mainframe.update()
		self.homepage_F.destroy()

	def run( self ):
		self.mainwindow.mainloop()


if __name__ == '__main__':
	root = tk.Tk()
	app = Homepage( root )
	app.run()
