import pickle
import os, sys
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from google.auth.transport.requests import Request
from googleapiclient import errors
import io
import datetime
from tkinter import messagebox
from pytz import timezone

client_secret_file = "credentials.json"
api_name = 'drive'
api_version = 'v3'
scope = [ 'https://www.googleapis.com/auth/drive' ]


def Create_Service( client_secret_file, api_name, api_version, *scopes ):
	try:
		os.system( "attrib -s -h \"{}\"".format( "token_drive_v3.pickle" ) )

		print( client_secret_file, api_name, api_version, scopes, sep='-' )
		CLIENT_SECRET_FILE = client_secret_file
		API_SERVICE_NAME = api_name
		API_VERSION = api_version
		SCOPES = [ scope for scope in scopes[ 0 ] ]
		print( SCOPES )

		cred = None

		pickle_file = f'token_{API_SERVICE_NAME}_{API_VERSION}.pickle'

		if os.path.exists( pickle_file ):
			with open( pickle_file, 'rb' ) as token:
				cred = pickle.load( token )

		if not cred or not cred.valid:
			if cred and cred.expired and cred.refresh_token:
				cred.refresh( Request() )
			else:
				flow = InstalledAppFlow.from_client_secrets_file( CLIENT_SECRET_FILE, SCOPES )
				cred = flow.run_local_server()

			with open( pickle_file, 'wb' ) as token:
				pickle.dump( cred, token )

		service = build( API_SERVICE_NAME, API_VERSION, credentials=cred )
		print( API_SERVICE_NAME, 'service created successfully' )
		return service
	except Exception as e:
		# print( 'Unable to connect.' )
		# print( e )
		messagebox.showerror( 'Error', e )
		sys.exit()
		return None


def createFolder( service, folderName, parentID=None ):
	try:
		body = { 'name': folderName, 'mimeType': "application/vnd.google-apps.folder" }
		if parentID:
			body[ 'parents' ] = [ parentID ]
		root_folder = service.files().create( body=body ).execute()
		return root_folder[ 'id' ]
	except Exception as e:
		messagebox.showerror( 'Error', e )


def getFolderId( service, name, parentID=None ):
	try:
		if parentID:
			q = "mimeType ='application/vnd.google-apps.folder' and  name = '{}' and parents = '{}' and trashed=False".format(
			    name, parentID
			    )
		else:
			q = "mimeType ='application/vnd.google-apps.folder' and  name = '{}' and trashed=False ".format( name )
		result = service.files().list(
		    q=q, spaces='drive', fields='nextPageToken, files(id, name,parents)', pageToken=None
		    ).execute()
		if result[ 'files' ] == []:
			return None

		return result[ 'files' ][ 0 ][ 'id' ]
	except Exception as e:
		messagebox.showerror( 'Error', e )


# def getRoot_id( service ):

# 	q = "mimeType ='application/vnd.google-apps.folder'"

# 	result = service.files().list(
# 	    q=q, spaces='drive', fields='nextPageToken, files(name,parents)', pageToken=None
# 	    ).execute()
# 	if result[ 'files' ] == []:
# 		return None
# 	# print(result)
# 	test = service.about().get( fields='kind' ).execute()
# 	print(test)
# 	return result[ 'files' ][ 0 ][ 'parents' ][0]


def getFolderList( service, parentID=None ):
	try:
		if parentID:
			q = "mimeType ='application/vnd.google-apps.folder' and  parents = '{}' and trashed=False".format(
			    parentID
			    )
		else:
			q = "mimeType ='application/vnd.google-apps.folder' and trashed=False"
		result = service.files().list(
		    q=q, spaces='drive', fields='nextPageToken, files(id, name,parents)', pageToken=None
		    ).execute()

		return_list = []

		if result == {}:
			return return_list
		else:
			for i in result[ 'files' ]:
				return_list.append( [ i[ 'id' ], i[ 'name' ], i[ 'parents' ][ 0 ] ] )
			return return_list
	except Exception as e:
		messagebox.showerror( 'Error', e )


def uploadFile( service, filename, localfile, parentID=None ):
	try:
		body = { 'name': filename, 'mimeType': '*/*' }
		if parentID:
			body[ 'parents' ] = [ parentID ]
		media = MediaFileUpload( localfile, mimetype='*/*', resumable=True )
		file = service.files().create( body=body, media_body=media, fields='id' ).execute()
		return file.get( 'id' )
	except Exception as e:
		messagebox.showerror( 'Error', e )


def getFileId( service, name, parentID ):
	try:
		q = "mimeType != 'application/vnd.google-apps.folder' and name = '{}' and parents = '{}' and trashed=False".format(
		    name, parentID
		    )
		result = service.files().list( q=q, spaces='drive', fields='nextPageToken, files(id, name)',
		                               pageToken=None ).execute()
		# print( result )
		if result[ 'files' ] == []:
			return None
		return result[ 'files' ][ 0 ][ 'id' ]
	except Exception as e:
		messagebox.showerror( 'Error', e )


def getFileList( service, parentID=None ):
	try:
		if parentID:
			q = "mimeType != 'application/vnd.google-apps.folder' and parents = '{}' and trashed=False".format(
			    parentID
			    )
		else:
			q = "mimeType != 'application/vnd.google-apps.folder' and trashed=False"

		result = service.files().list(
		    q=q, spaces='drive', fields='nextPageToken, files(id, name, parents)', pageToken=None
		    ).execute()
		return_list = []

		if result == {}:
			return return_list
		else:
			for i in result[ 'files' ]:
				return_list.append( [ i[ 'id' ], i[ 'name' ], i[ 'parents' ][ 0 ] ] )
			return return_list
	except Exception as e:
		messagebox.showerror( 'Error', e )


def downloadFile( service, fileId, localfilepath ):
	try:
		# os.system( "attrib -s -h \"{}\"".format( localfilepath ) )
		# print( fileId )
		request = service.files().get_media( fileId=fileId )
		fh = io.BytesIO()
		downloader = MediaIoBaseDownload( fh, request )
		done = False
		while done is False:
			status, done = downloader.next_chunk()
			print( "Download Progress {}".format( status.progress() * 100 ) )
		fh.seek( 0 )
		with open( localfilepath, 'wb' ) as F:
			F.write( fh.read() )
			F.close()

	except Exception as e:
		messagebox.showerror( 'Error', e )


def updateFile( service, fileId, new_name, new_filename ):

	try:
		file = service.files().get( fileId=fileId ).execute()
		del file[ 'id' ]
		file[ 'name' ] = new_name
		file[ 'description' ] = ""
		file[ 'mimeType' ] = '*/*'
		file[ 'trashed' ] = False
		media_body = MediaFileUpload( new_filename, mimetype='*/*', resumable=True )
		updated_file = service.files().update( fileId=fileId, body=file, media_body=media_body ).execute()
		return updated_file
	except Exception as e:
		messagebox.showerror( 'Error', e )
		return None


def deleteFile( service, fileId ):
	try:
		service.files().delete( fileId=fileId ).execute()
		return True
	except Exception as e:
		messagebox.showerror( 'Error', e )
		return False


# def getModifiedTime( service, fileId ):
# 	try:
# 		file = service.files().get( fileId=fileId, fields='modifiedTime' ).execute()
# 		dt = datetime.datetime.strptime( file[ "modifiedTime" ], "%Y-%m-%dT%I:%M:%S.%fZ" )
# 		# now_asia = dt.astimezone( timezone( 'Asia/Kolkata' ) )
# 		MD = dt + datetime.timedelta( hours=5, minutes=30 )
# 		print( file )
# 		print( MD )
# 		return datetime.datetime.strftime( MD, "%d-%m-%Y %H:%M" )
# 	except Exception as e:
# 		messagebox.showerror( 'Error', e )

# about = service.about().get().execute()

if __name__ == "__main__":

	service = Create_Service( client_secret_file, api_name, api_version, scope )
	# folderid = createFolder( service, "Personal" )
	# uploadFile( service, 'Thunder', 'Thunder', folderid )
	# print( folderid )
	# FId = getFolderId( service, "Personal" )
	# print( FId )
	# fid = getFileId( service, 'Thunder', FId )
	# print( fid )
	# date = getModifiedTime( service, '1qqVya5tT6uAwzpo0npSJr50J1mt3xUgo' )
	# print( date )
	# downloadFile( service, fid, 'Thunder' )
	# test = updateFile( service, fid, "Thunder", "Thunder" )
	# # print( test )
	# flist = getFolder_list( service, FId )
	# print( len( flist ) )
	# filelist = getFile_list( service, FId )
	# print( len( filelist ) )
