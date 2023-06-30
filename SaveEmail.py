#Email save

import win32com.client
import os

pathtosave = r'sample path'

outlook = win32com.client.Dispatch('Outlook.Application')
mapi = outlook.GetNamespace('MAPI')
inbox = mapi.GetDefaultFolder(6)
messages = inbox.items

samplesubject = 'sample subject text'

for msg in messages:
	if msg.subject == samplesubject:
		msgname = msg.subject
		msgname = str(msgname)
		print(msgname)
		msg.SaveAs(os.path.join(pathtosave, (msg.Subject + '.msg')))