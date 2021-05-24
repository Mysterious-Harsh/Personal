import os
from cryptography.fernet import Fernet
# import pickle
# from win32ui import MessageBox
from tkinter import messagebox
import shutil


class check_registration:

	def check( self ):
		self.home = os.path.abspath( os.path.expanduser( '~' ) )
		self.filepath = os.path.join( self.home, "Invisible_Key", "do_not_delete.key.enc" )
		self.pwd = os.path.abspath( os.path.dirname( __file__ ) )
		self.filepath2 = os.path.join( self.pwd, "Invisible_Key", "registration_key.key.shp" )
		self.key1 = "pAvLFxFMGiqHbri3a2N4XThSnicwZJd7i5G9LVvfv_Q="
		self.key2 = "HAF_rZF-CioBeUnhFIWFa4_K-7Jwf4qMLOJmDHdjssM="
		self.key3 = "xeifEwZoy18YNOGAp_ZyaPgidgTwtTvDHkunw0SOmlU="
		try:

			if os.path.exists( self.filepath ) and os.path.exists( self.filepath2 ):
				f1 = Fernet( self.key1 )
				f2 = Fernet( self.key2 )
				f3 = Fernet( self.key3 )

				with open( self.filepath, "rb" ) as EF:
					encrypted_data = EF.read()

				decrypted_data1 = f3.decrypt( encrypted_data )
				decrypted_data2 = f2.decrypt( decrypted_data1 )
				match1 = f1.decrypt( decrypted_data2 ).decode()
				# print( decrypted_data )

				self.key4 = "l1My-18hcmH_tyFkvRh_Wcc3UHN7yE3jHW25OW4i2Lw="
				self.key5 = "sRkfpJcgi0nkXj3L0RfaEqoSwrF0iqToCwecAMACGvk="
				self.key6 = "d0S9tNfKo0bIO8g33_o8n09WP1PtDmx2OENAUEq6HuU="

				f4 = Fernet( self.key4 )
				f5 = Fernet( self.key5 )
				f6 = Fernet( self.key6 )

				with open( self.filepath2, "rb" ) as EF:
					# encrypted_data = pickle.load( EF )
					encrypted_data = EF.read()

				decrypted_data1 = f6.decrypt( encrypted_data )
				decrypted_data2 = f5.decrypt( decrypted_data1 )
				match2 = f4.decrypt( decrypted_data2 ).decode()

				# MAC = gma()
				if match1 == match2:
					return True
				else:
					return False
		except:
			pass
		return False

	def Register( self, registration_key ):
		self.home = os.path.abspath( os.path.expanduser( '~' ) )
		self.filepath = os.path.join( self.home, "Invisible_Key", "do_not_delete.key.enc" )
		DIR = os.path.join( self.home, "Invisible_Key" )
		try:
			shutil.rmtree( DIR )
		except:
			pass
		try:
			os.makedirs( DIR )
		except:
			pass
		# filename, key = os.path.split( self.filepath )
		self.key1 = "pAvLFxFMGiqHbri3a2N4XThSnicwZJd7i5G9LVvfv_Q="
		self.key2 = "HAF_rZF-CioBeUnhFIWFa4_K-7Jwf4qMLOJmDHdjssM="
		self.key3 = "xeifEwZoy18YNOGAp_ZyaPgidgTwtTvDHkunw0SOmlU="

		f1 = Fernet( self.key1 )
		f2 = Fernet( self.key2 )
		f3 = Fernet( self.key3 )

		file_data = registration_key.encode()

		encrypted_data1 = f1.encrypt( file_data )
		encrypted_data2 = f2.encrypt( encrypted_data1 )
		encrypted_data = f3.encrypt( encrypted_data2 )

		with open( self.filepath, "wb" ) as f:
			f.write( encrypted_data )
			f.flush()
			# pickle.dump( encrypted_data, f )

		os.system( "attrib +s +h \"{}\\*.*\" /s".format( DIR ) )
		os.system( "attrib +s +h \"{}\"".format( DIR ) )

		self.pwd = os.path.abspath( os.path.dirname( __file__ ) )
		self.filepath = os.path.join( self.pwd, "Invisible_Key", "registration_key.key.shp" )
		DIR = os.path.join( self.pwd, "Invisible_Key" )
		try:
			shutil.rmtree( DIR )
		except:
			pass
		try:
			os.makedirs( DIR )
		except:
			pass
		# filename, key = os.path.split( self.filepath )
		self.key4 = "l1My-18hcmH_tyFkvRh_Wcc3UHN7yE3jHW25OW4i2Lw="
		self.key5 = "sRkfpJcgi0nkXj3L0RfaEqoSwrF0iqToCwecAMACGvk="
		self.key6 = "d0S9tNfKo0bIO8g33_o8n09WP1PtDmx2OENAUEq6HuU="

		f4 = Fernet( self.key4 )
		f5 = Fernet( self.key5 )
		f6 = Fernet( self.key6 )

		file_data = registration_key.encode()

		encrypted_data1 = f4.encrypt( file_data )
		encrypted_data2 = f5.encrypt( encrypted_data1 )
		encrypted_data = f6.encrypt( encrypted_data2 )

		with open( self.filepath, "wb" ) as f:
			f.write( encrypted_data )
			f.flush()
			# pickle.dump( encrypted_data, f )

		os.system( "attrib +s +h \"{}\\*.*\" /s".format( DIR ) )
		os.system( "attrib +s +h \"{}\"".format( DIR ) )

		messagebox.showinfo( "Success", "Congratulation, Your PC is registerd." )


# t = check_registration()
# t.Register( "FYhy6TaRgJY6ftQe7imAqpj2" )
# print( t.check() )
