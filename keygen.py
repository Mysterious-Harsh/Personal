import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
import hashlib


class KeygenApp:

	def __init__( self, master=None ):
		# build ui
		self.root = master
		self.keygen_F = ttk.Labelframe( master )
		self.keygen_E = ttk.Entry( self.keygen_F )
		self.keygen_E.config( font='{Arial} 12 {bold}', takefocus=True, width='32' )
		self.keygen_E.grid()
		self.keygen_B = ttk.Button( self.keygen_F )
		self.keygen_B.config( default='normal', takefocus=True, text='Generate' )
		self.keygen_B.grid( column='1', row='0' )
		self.keygen_B.configure( command=self.generate )
		self.keygen_L = ttk.Entry( self.keygen_F )
		self.keygen_LTV = tk.StringVar()
		self.keygen_L.config( font='{Arial} 12 {bold}', textvariable=self.keygen_LTV, width='45' )
		self.keygen_L.grid( column='0', columnspan='2', row='1' )
		self.keygen_F.config( height='200', labelanchor='n', relief='flat', text='Key Gen' )
		self.keygen_F.config( width='200' )
		self.keygen_F.pack( padx='5', pady='5', side='top' )
		self.keygen_E.focus_set()
		# Main widget
		self.mainwindow = self.keygen_F
		self.config()

	def config( self ):
		self.filepath = ""
		# self.key = "0ZxFqrHjpmY2LBbL-bUiQfvsr0DV1kCBBQEbbsZdRWU="
		self.s_w = self.mainwindow.winfo_screenwidth()
		self.s_h = self.mainwindow.winfo_screenheight()
		self.wow = 425
		self.how = 100
		x_c = ( ( self.s_w / 2 ) - ( self.wow / 2 ) )
		y_c = ( ( self.s_h / 2 ) - ( self.how / 2 ) )
		self.root.geometry( "%dx%d+%d+%d" % ( self.wow, self.how, x_c, y_c ) )
		self.root.resizable( False, False )
		self.style = ttk.Style( self.mainwindow )
		# ##print( self.style.theme_names() )
		# self.style.theme_use( "winnative" )
		self.style.configure( "TLabelframe", background="#ffffff", highlightthickness=5, relief="flat" )
		self.style.configure( "TFrame", background="#44ccff", highlightthickness=5, relief="flat" )
		self.style.configure(
		    "TLabelframe.Label",
		    background="#081944",
		    foreground="#ffffff",
		    font=( 'Arial', 14, 'bold' ),
		    width=self.wow,
		    anchor="center",
		    relief="flat"
		    )
		self.style.configure(
		    "Treeview",
		    fieldbackground="#ffffff",
		    background="#ffffff",
		    foreground="#000000",
		    font=( "Arial", 11, "bold" ),
		    borderwidth=5,
		    relief="flat"
		    )
		self.style.configure( "Treeview.Heading", font=( "Arial", 12, "bold" ) )

		self.style.configure(
		    "TButton",
		    background="#ffffff",
		    foreground="#000000",
		    font=( "Arial", 11, "bold" ),
		    relief="flat",
		    highlightthickness=5,
		    padding=5,
		    highlightcolor="blue",
		    bordercolor="#000000",
		    takefocus=True,
		    width=12
		    )
		# self.style.map( "TButton", background=[ ( 'active', 'blue' ) ] )
		self.style.configure(
		    "TLabel",
		    padding=0,
		    font=( "Arial", 11, "bold" ),
		    relief="flat",
		    width=14,
		    takefocus=False,
		    borderwidth=0,
		    cursor="arrow",
		    anchor="center"
		    )

		self.style.configure( "TEntry", fieldbackground="#99ddff", font=( "Arial", 11, "bold" ) )
		# self.filepath_E.config( state="readonly" )

	def add_attachment( self, event=None ):
		self.filepath = filedialog.askopenfilename(
		    initialdir="/home/blackthunder/Thunder/My Projects/restriction/",
		    title="Select file",
		    filetypes=( ( "All Files", "*.*" ), ( "files", ".doc* .enc" ), ( "enc file", "*.enc" ) )
		    )

	def generate( self ):
		key_str = self.keygen_E.get()
		if key_str == "":
			messagebox.showerror( "Empty", "Empty Field !" )
		else:
			temp = []
			hash_key = hashlib.md5(
			    hashlib.sha1(
			        hashlib.sha256( hashlib.sha512( key_str.encode() ).hexdigest().encode() ).hexdigest().encode()
			        ).hexdigest().encode()
			    ).hexdigest()
			for i, j in zip( [ 0, 8, 16, 24 ], [ 8, 16, 24, 32 ] ):
				temp.append( hash_key[ i : j ] )
			print( temp )
			# print(len(hash_key))
			self.keygen_LTV.set( '-'.join( temp ) )

		pass

	def run( self ):
		self.mainwindow.mainloop()


if __name__ == '__main__':
	root = tk.Tk()
	app = KeygenApp( root )
	app.run()
