import win32com.client
import datetime as dt
from pathlib import Path
import shutil
import os
import fnmatch
from subprocess import call

WorkDir = "D:\\attachments"
FolderName = "Remote"
Pscp = "C:\\Program Files (x86)\\PuTTY\\pscp.exe"
Subject = "Archive Attachments"

FolderOutput = WorkDir + "\\files"
LockDir = WorkDir + "\\locks"
ZipDir = WorkDir + "\\zip"
HostsFile = os.path.dirname(os.path.realpath(__file__)) + "\\remote_hosts.txt"
EmailsFile = os.path.dirname(os.path.realpath(__file__)) + "\\emails.txt"
LogDir = WorkDir + "\\logs"

if not os.path.exists(FolderOutput):
    os.makedirs(FolderOutput)

if not os.path.exists(ZipDir):
    os.makedirs(ZipDir)

if not os.path.exists(LockDir):
    os.makedirs(LockDir)

if not os.path.exists(LogDir):
    os.makedirs(LogDir)

LogFileName = LogDir + "\\" + dt.datetime.today().strftime("%Y-%m-%d") + ".log"
if (Path(LogFileName).is_file() == False):
    LogFile = open(LogFileName, "w+")

LogFile = open(LogFileName, "a")

def Log(msg):
    LogFile.write(dt.datetime.today().strftime("%Y-%m-%d %H:%M:%S") + " " + msg + "\n")

Log("Start script")

app =  win32com.client.Dispatch('Outlook.Application')
outlook = app.GetNamespace('MAPI')

if (not FolderName):
    inbox = outlook.GetDefaultFolder(6)
else:
    inbox = outlook.GetDefaultFolder(6).Folders(FolderName) 

messages = inbox.Items
for message in messages:
    if message.UnRead:
        Log("Open Message with subject: " + message.Subject)
        for attachment in message.attachments:
            Log("Message has attachment: " + attachment.FileName)
            if (fnmatch.fnmatch(attachment.FileName, '*.csv') or fnmatch.fnmatch(attachment.FileName, '*.txt')):
                Log("Copy attachment " + attachment.FileName + " to " + FolderOutput)
                attachment.SaveAsFile(FolderOutput + '\\' + attachment.FileName)
                message.UnRead = False

lock = LockDir + "\\" + dt.datetime.today().strftime("%Y-%m-%d-%H")
output_filename = ZipDir + "\\" + dt.datetime.today().strftime("%Y-%m-%d-%H-%M-%S") + ".zip"

print(os.path.dirname(os.path.realpath(__file__)))

if (dt.datetime.today().hour == 23):
    if (Path(lock).is_file() == False):
        Log("Start archiving to zip and sending by emails")
        open(lock, "w+")
        shutil.make_archive(output_filename, 'zip', FolderOutput)
        Log("Files from dir: " + FolderOutput)
        for f in os.listdir(FolderOutput):
            Log(f)
        Log("Added to archive: " + output_filename)
        
        with open(HostsFile, "r") as lines:
            for line in lines:
                command = '"' + Pscp + '"' + " -pw " + line.split(';')[2] + " " + output_filename + ".zip " + line.split(';')[1] + "@" + line.split(';')[0] + ":/tmp"
                os.system(command)
                Log("Remote copy " + output_filename + " to server " + line.split(';')[0])

        with open(EmailsFile, "r") as lines:
            for line in lines:
                Log("Send " + output_filename + ".zip by email " + line)
                Email = app.CreateItem(0)
                Email.To = line
                Email.Subject = Subject
                Email.Attachments.Add(output_filename + ".zip")
                Email.Send
                Email.Move(outlook.GetDefaultFolder(4))
        
#        shutil.rmtree(ZipDir)
        if not os.path.exists(ZipDir):
            os.makedirs(ZipDir)

Log("End script")
LogFile.close()