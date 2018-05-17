# python 3
# Install Third party module using "pip install <module-name>" - for windows


# IMPORTING MODULES

import email
import imaplib
from validate_email import validate_email
import getpass
import os
import sys
import mimetypes
import subprocess
import shutil
import time
from pathlib import Path
import easygui
import datetime, schedule
from threading import Thread


# MAIN MENU
# TO SELECT WHETHER TO SEND EMAIL, VIEW EMAIL OR SIGN OUT AND EXIT

def main_menu():

    while True:
        print("")
        print("----- MAIN MENU -------")
        print("1. SEND EMAIL MESSAGE(s)")
        print("2. VIEW EMAIL MESSAGES(s)")
        print("3. SIGN OUT AND EXIT")
        print("")
        print("ENTER THE CORRESPONDING INDEX OF YOUR CHOICE: ", end='')
        ch = input()
        if ch == '1':
            to=''
            sendemail()
        elif ch == '2':
            mail = imaplib.IMAP4_SSL("imap.gmail.com")     
            mail.login(username,pswd)
            view()
        elif ch == '3':
            mail = imaplib.IMAP4_SSL("imap.gmail.com")     
            mail.login(username,pswd)
            mail.logout()
            sys.exit()
        else:
            print('')
            print("INVALID INPUT! ENTER AGAIN.")
            print('')


# FUNCTION TO OPEN DIALOG BOX TO SELECT FILES

def dialog():
    path = easygui.fileopenbox()
    return path


# FUNCTION TO SEND EMAIL USING sendEmail API

def sendemail():
    a=[os.path.join(os.getcwd(),'sendEmail.exe')]
    a.append('-f')
    a.append(username)
    a.append('-xu')
    a.append(username)
    a.append('-xp')
    a.append(pswd)
    a.append('-s')
    a.append('smtp.gmail.com:587')
    
    a.append('-t')
    while True:
        print('')
        print("ENTER RECIPIENT'S EMAIL ADDRESS  ", end='')
        a.append(input())
        print("ADD MORE RECIPIENTS? (y/n): ", end='')
        if input() == 'n':
            break

    a.append('-u')
    print('')
    print("ENTER SUBJECT: ", end='')
    a.append(input())
    a.append('-m')
    print('')
    print("ENTER MESSAGE: ",)
    print('')
    print("WRITE YOUR MESSAGE. SAVE IT AND CLOSE THE EDITOR.")
    time.sleep(1)
    Path('mail-body.txt').touch()
    p = os.path.join(os.getcwd(), 'mail-body.txt')
    os.system(p)
    with open("mail-body.txt") as f:
        file_content = f.read().rstrip("\n")
    a.append(file_content)
    os.remove("mail-body.txt")
    print('')
    print("MESSAGE BODY:")
    print('')
    print(file_content)
    print('')
    print("ADD ATTACHMENTS? (y/n): ", end='')
    if input() == 'y':
        a.append('-a')
        while True:
            path = dialog()
            a.append(path)
            print('')
            print("ADD MORE ATTACHMENTS? (y/n): ", end='')
            if input() == 'n':
                break
    print('')
    print("ADD CC RECIPIENTS? (y/n): ", end='')
    if input() == 'y':
        a.append('-cc')
        while True:
            print('')
            print("ENTER RECIPIENT'S EMAIL ADDRESS: ", end='')
            a.append(input())
            print("ADD MORE RECIPIENTS? (y/n): ", end='')
            if input() == 'n':
                break
    print('')
    print("ADD BCC RECIPIENTS? (y/n): ",end='')
    if input() == 'y':
        a.append('-bcc')
        while True:
            print('')
            print("ENTER RECIPIENT'S EMAIL ADDRESS: ", end='')
            a.append(input())
            print("ADD MORE RECIPIENTS? (y/n): ", end='')
            if input() == 'n':
                break

    while True:
        print("1. SEND EMAIL NOW")
        print("2. SEND EMAIL LATER")
        print("")
        print("ENTER THE CORRESPONDING INDEX OF YOUR CHOICE: ", end='')
        ch2 = input()
        if ch2 =='1':
# EMAIL IS SENT
            print('')
            print("SENDING EMAIL ...")
            subprocess.call(a)
            main_menu()
            break
        elif ch2 == '2':
            b=a
            
            # SCHEDULING OF EMAIL


            print("ENTER DATE IN THE FOLLOWING FORMAT: DD-MM-YYYY")
            print("ENTER DATE: ", end='')
            d= input()
            print("ENTER TIME IN THE FOLLOWING FORMAT: HH:MM:SS")
            print("ENTER TIME: ", end='')
            t=input()
            dt = d + " " + t
            print("")
            print("YOUR EMAIL WILL BE SENT ON: "+ dt)
            print("")
            def job2():
                def job():

                    date = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                    if date == dt:
                        subprocess.call(b)
                        sys.exit()
                schedule.every(0.01).minutes.do(job)

                while True:
                    schedule.run_pending()
                    time.sleep(1)

            th= Thread(target = job2)
            th2 = Thread(target = main_menu)
            
            th.start()
            th2.start()
            sys.exit()

        else:
            print('')
            print("INVALID INPUT! ENTER AGAIN.")
            print('')

# FUNCTION TO FETCH EMAILS IN A FOLDER


def fetch_email(length, j, label_name, inbox_item_list):
    
    for index in range(length-1,-1,-1):
        result2, email_data = mail.uid('fetch', inbox_item_list[index], '(RFC822)')
        raw_email = email_data[0][1].decode("utf-8")
        email_message = email.message_from_string(raw_email)
        print('')
        print(str(j) + '.')
        print("Date: " + email_message['Date'])
        print("To: " + email_message['To'])
        print("From: " + email_message['From'])
        print("Subject: " + email_message['Subject'])
        print("\n")
        j=j+1
        if j%6==0 or index == 0:
            print('')
            print("1. VIEW OLDER")
            print("2. VIEW FULL EMAIL MESSAGE")
            print("3. VIEW EMAIL FROM OTHER LABELS")
            print("4. MAIN MENU")
            print("5. SIGN OUT AND EXIT")
            print("ENTER CORRESPONDING INDEX OF YOUR CHOICE: ", end='')
            c= input()
            if c == '1':
                if index == 0:
                    print('')
                    print('NO MORE EMAILS TO VIEW')
                    while True:
                        print('')
                        print('1. BACK TO ' + label_name)
                        print("2. LABEL SELECTION MENU")
                        print('')
                        print('ENTER THE CORRESPONDING INDEX OF YOUR CHOICE: ', end='')
                        i = input()
                        if i == '1':
                            fetch_email(length, 1, label_name, inbox_item_list)
                        elif i == '2':
                            view()
                        else:
                            print('')
                            print('INVALID INPUT! ENTER AGAIN.')
                            print('')
                    
                else:
                    continue
                
            elif c == '2':
                print("ENTER THE CORRESPONDING INDEX OF THE EMAIL YOU WANT TO VIEW: ", end='')
                ind = int(input())
                detail_email(inbox_item_list, ind, mail, label_name,length)
            elif c == '3':
                view()
            elif c == '4':
                main_menu()
            else:
                mail.logout()
                sys.exit()



# FUNCTION TO SELECT WHICH EMAILS TO VIEW

def view():
    

    while True:
        print('')
        print("1. ALL MAIL")
        print("2. INBOX")
        print("3. DRAFTS")
        print("4. IMPORTANT")
        print("5. SENT MAIL")
        print("6. SPAM")
        print("7. STARRED")
        print("8. TRASH")
        print("9. CUSTOM LABEL CREATED")
        print("10. SIGN OUT AND EXIT")
        print("11. MAIN MENU")
        print("12. SEARCH EMAIL MESSAGE(s)")
        print("ENTER THE CORRESPONDING INDEX OF YOUR LABEL CHOICE: ", end='')
        label_choice=input()
        if label_choice == '1':
            label = '"[Gmail]/All Mail"'
            label_name= "ALL MAIL"
            break        
        
        if label_choice == '2':
            label = "INBOX"
            label_name= "INBOX"
            break
        elif label_choice == '3':
            label = "[Gmail]/Drafts"
            label_name= "DRAFTS"
            break
        elif label_choice == '4':
            label = "[Gmail]/Important"
            label_name= "IMPORTANT"
            break
        elif label_choice == '5':
            label = '"[Gmail]/Sent Mail"'
            label_name= "SENT MAIL"
            break
        elif label_choice == '6':
            label = "[Gmail]/Spam"
            label_name= "SPAM"
            break
        elif label_choice == '7':
            label = "[Gmail]/Starred"
            label_name= "STARRED"
            break
        elif label_choice == '8':
            label = "[Gmail]/Trash"
            label_name= "TRASH"
            break
        elif label_choice == '9':
            print("ENTER THE LABEL:", end='')
            label=input()
            label_name= label
            break
        elif label_choice =='10':
            mail.logout()
            sys.exit()
        elif label_choice == '11':
            main_menu()
        elif label_choice == "12":
            mail.select('"[Gmail]/All Mail"')
            label_name= "ALL MAIL"
            while True:
                print("")
                print("SEARCH USING")
                print("1. FROM")
                print("2. TO")
                print("3. SUBJECT")
                print("")
                print("ENTER CORRESPONDING INDEX OF YOUR CHOICE: ",end='')
                choi = input()
                if choi == '1':
                    print("")
                    print("ENTER EMAIL ADDRESS: ", end='')
                    srh = "FROM " + input()
                    
                    result, data = mail.uid('search', None, srh)
                    break
                elif choi == '2':
                    print("")
                    print("ENTER EMAIL ADDRESS: ", end='')
                    srh = "TO " + input()
                    result, data = mail.uid('search', None, srh)
                    break
                elif choi == '3':
                    print("")
                    print("ENTER SUBJECT LINE: ")
                    srh = 'SUBJECT "' + input() + '"'
                    result, data = mail.uid('search', None, srh)
                    break
                else:
                    print("")
                    print("INVALID INPUT! ENTER AGAIN")
                    print("")
                

            inbox_item_list = data[0].split()

            length = len(inbox_item_list)
            if length == 0:
                print('')
                print("NO MAILS")
                view()
            else:
                j=1
                print('')
          
                fetch_email(length, j, label_name, inbox_item_list)
            
        else:
            print('')
            print("INVALID INPUT! ENTER AGAIN.")
            print('')



    mail.select(label)

    while True:
        print("")
        print("1. VIEW ALL EMAILS IN " + label_name)
        print("2. VIEW EMAILS AFTER A SPECIFIC DATE")
        print("3. VIEW EMAILS BEFORE A SPECIFIC DATE")
        print("4. VIEW EMAILS AT A SPECIFIC DATE")
        print("5. UNSEEN MESSAGES")
        print("6. SEEN MESSAGES")
        print("")
        print("ENTER COORESPONDING INDEX OF YOUR CHOICE: ", end='')
        cho = input()
        if cho == '1':
            srch = "ALL"
            break
        elif cho == '2':
            print("ENTER THE DATE IN FORMAT ->  05-JUL-2015  :  ")
            srch = "SINCE " + input()
            break
        elif cho == '3':
            print("ENTER THE DATE IN FORMAT ->  05-JUL-2015  :  ")
            srch = "BEFORE " + input()
            break
        elif cho == '4':
            print("ENTER THE DATE IN FORMAT ->  05-JUL-2015  :  ")
            srch = "ON " + input()
            break
        elif cho == '5':
            srch = "UNSEEN"
            break
        elif cho == '6':
            srch = "SEEN"
            break
            
        else:
            print('')
            print("INVALID INPUT! ENTER AGAIN.")
            print('')

    result, data = mail.uid('search', None, srch)

    inbox_item_list = data[0].split()

    length = len(inbox_item_list)
    if length == 0:
        print('')
        print("NO MAILS IN " + label_name)
        view()
    else:
        j=1
        print('')
          
        fetch_email(length, j, label_name, inbox_item_list)                


# FUNCTION TO VIEW FULL EMAIL MESSAGE

def detail_email(inbox_item_list, index, mail, label_name, length):
    result2, email_data = mail.uid('fetch', inbox_item_list[-index], '(RFC822)')
    raw_email = email_data[0][1].decode("utf-8")
    email_message = email.message_from_string(raw_email)
    print('')
    print("Date: " + email_message['Date'])
    print("To: " + email_message['To'])
    print("From: " + email_message['From'])
    print("Subject: " + email_message['Subject'])


# GIVING FILE EXTENSION TO MESSAGE BODY
    counter = 1

    for part in email_message.walk():
        if part.get_content_maintype() == "multipart":
            continue
        filename = part.get_filename()
        content_type = part.get_content_type()
        if not filename:
            ext = mimetypes.guess_extension(content_type)
            if not ext:
                ext = '.bin'
            if 'plain' in content_type:
                ext = '.txt'
            
            filename = 'msg-part-%08d%s' %(counter, ext)
        counter+=1

# SAVING THE MESSAGE AND OPENING THE FILE TO VIEW MESSAGE BODY
    print("Content Type: " + content_type)

    # VIEW PLAIN TEXT EMAIL MESSAGE
    if "plain" in content_type:
        print('')
        save_path= os.path.join(os.getcwd(),'openemails')
        if not os.path.exists(save_path):
            os.makedirs(save_path)
            
        with open(os.path.join(save_path,filename),'wb') as fp:
            fp.write(part.get_payload(decode=True))
        print('')
        input("PRESS ENTER TO VIEW MESSAGE BODY:")
        
        os.startfile(os.path.join(save_path, filename))
    # VIEW HTML EMAIL MESSAGE
    elif "html" in content_type:
        print('')
        save_path= os.path.join(os.getcwd(),'openemails')
        if not os.path.exists(save_path):
            os.makedirs(save_path)
            
        with open(os.path.join(save_path,filename),'wb') as fp:
            fp.write(part.get_payload(decode=True))
        print('')
        input("PRESS ENTER TO VIEW MESSAGE BODY:")
        
        os.startfile(os.path.join(save_path, filename))
        
# ATTACHMENT DOWNLOAD        
    else:
        while True:
            print('')
            print("DOWNLOAD ATTACHMENTS? (y/n): ", end='')
            choice = input()
            if choice == 'n':
                break
            elif choice == 'y':
                
                save_path=r'C:\Users\Rishav\Downloads'
                if not os.path.exists(save_path):
                    os.makedirs(save_path)
            
                with open(os.path.join(save_path,filename),'wb') as fp:
                    fp.write(part.get_payload(decode=True))
                    print('')
                print("YOUR ATTACHMENT IS SAVED IN: " + str(save_path))
                break

            else:
                print('')
                print("INVALID INPUT! ENTER AGAIN.")
                print('')
    
    while True:
        
        print('')
        print("1. SAVE EMAIL MESSAGE TO THIS PC")
        print("2. BACK TO LABEL SELECTION TO VIEW EMAILS")
        print("3. BACK TO " + label_name + " EMAILS")
        print("4. MAIN MENU")
        print("5. SIGN OUT AND EXIT")
        print("ENTER CORRESPONDING INDEX OF YOUR CHOICE:", end='')
        c2 = input()
   
        if c2 == '1':
            # SAVE EMAIL MESSAGE AND ATTACHMENTS
           save_path= os.path.join(r'C:\Users\Rishav\Downloads','emails')
           if not os.path.exists(save_path):
               os.makedirs(save_path)
            
           with open(os.path.join(save_path,filename),'wb') as fp:
               fp.write(part.get_payload(decode=True))
           print('')
           print("EMAIL MESSAGE SAVED.")
           print("YOUR SAVED EMAIL MESSAGE IS IN LOCATION: "+ save_path)
        elif c2 == '2':
            view()
        elif c2 == '3':
            fetch_email(length,1,label_name, inbox_item_list)
        elif c2 == '4':
            main_menu()
        elif c2 == '5':
            mail.logout()
            sys.exit()
        else:
            print('')
            print("INVALID INPUT! ENTER AGAIN.")
            print('')

# PROGRAM START


# REMOVING THE FOLDER WHICH CONTAINS FILES WHICH ARE USED TO OPEN MESSAGE BODY        
if os.path.exists(os.path.join(os.getcwd(), 'openemails')):
    shutil.rmtree(os.path.join(os.getcwd(), 'openemails'))




# USER LOGIN AND AUTHENTICATION

while True:
    print("---SIGN IN TO YOUR EMAIL ACCOUNT----")
    print("\n")
    try:
        s = "ENTER YOUR EMAIL ADDRESS: "
        while True:
            
            print (s, end='')
            username= input()
            is_valid = validate_email(username)
            if is_valid == True:
                break
            s = "ENTER A VALID EMAIL ADDRESS: "

        pswd = getpass.getpass('PASSWORD: ')
        mail = imaplib.IMAP4_SSL("imap.gmail.com")     
        mail.login(username,pswd)
        main_menu()
    
    
    except imaplib.IMAP4.error:
        print("YOUR EMAIL ID AND PASSWORD DID NOT MATCH. LOGIN FAILED! ENTER YOUR CREDENTIALS AGAIN.")
       


