'''
Created on Aug 7, 2015
@author: Dar1
'''
import smtplib
from email.mime.text import MIMEText
from email.MIMEMultipart import MIMEMultipart
from email.MIMEImage import MIMEImage
from pandas import *
import getpass
from datetime import datetime, timedelta
import re
import numpy as np
import os
import xlrd
try:
    from AcademyProcessFig import build_progress_bar
except:
    pass
from email.mime.application import MIMEApplication
import math

xl_format = "-------------------------------------------------------------------------------------------------------------------\n\
\nHow the track_contacts.xlsx should look like? \n\n\
Timestamp       ||FullName      ||email                   ||track      ||StartEnd            ||branch    \n\
=========================================================================================================\n\
5/8/2015        ||Dar Lador     ||dar.lador@gmail.com     ||web        ||Start_8.15_End_3.16 ||WIS        \n\
5/8/2015        ||Foo Bar       ||foo.bar@gmail.com       ||android    ||Start_8.15_End_3.16 ||WIS        \n\
5/8/2015        ||Debugi Bugi   ||debugi.bugi@outlook.com ||short_web  ||Start_8.15_End_9.15 ||WIS        \n\
\nThe program doesn't care what is written in the columns: Timestamp, StartEnd and branch.\n\
However: \n(1) FullName must be written in English (case-insensitivity) \n(2) email address must be written exactly as shown\n\
(3) track must be written as 'web', 'android' or 'short_meb'.\n(4) The sheet name is 'tracks'\n\
-------------------------------------------------------------------------------------------------------------------\n"

shecodes = "\n\
=========================================================================================================================\n\
=========================================================================================================================\n\
\n\
  SSSSSSSS     HH    HH   EEEEEEEE            CCCCCCCC     OOOOOOOOOO       DDDDDDD       EEEEEEEE     SSSSSSSS      OOO\n\
 SSS           HH    HH   EE                CCCC         OOOO      OOOO     DD    DDD     EE          SSS           OOOOO\n\
SS             HH    HH   EE               CCC          OOO          OOO    DD      DDD   EE         SS              OOO\n\
 SSS           HHHHHHHH   EEEEEEEE        CC           OO              OO   DD        DD  EEEEEEE     SSS                \n\
   SSSSSS      HHHHHHHH   EEEEEEEE        C            O                O   DD         D  EEEEEEE       SSSSS           ,\n\
       SSS     HH    HH   EE              CC           OO              OO   DD        DD  EE                SSS         ,, \n\
         SS    HH    HH   EE               CCC          OOO          OOO    DD      DDD   EE                  SS       ,,,\n\
       SSS     HH    HH   EE                CCCC         OOOO      OOOO     DD    DDD     EE                SSS      ,,,,\n\
 SSSSSSS       HH    HH   EEEEEEEE            CCCCCCCC     OOOOOOOOOO       DDDDDDD       EEEEEEEE    SSSSSSS      ,,,,,\n\
\n\
=========================================================================================================================\n\
========================================================================================================================="

error_ = "\n\
****************************************************************\n\
\n\
  EEEEE     RRRRRR       RRRRRR        OOOOOO        RRRRRR\n\
  EE        RR   RR      RR   RR     OOO    OOO      RR   RR\n\
  EERRR     RR   R       RR   R     OO        OO     RR   R\n\
  EEEEE     RRRRR        RRRRR      OO        OO     RRRRR\n\
  EE        RR   RR      RR   RR     OOO    OOO      RR   RR\n\
  EEEEE     RR   RRRR    RR   RRRR     OOOOOO        RR   RRRR\n\
\n\
****************************************************************\n\
\n\
"
#########################################################################################################################
#                                   Attend
#########################################################################################################################
class Attend():
    def __init__(self, error_, my_path, debug=None, members_update=False):
        self.error = error_
        self.my_path = my_path
        self.w = read_excel(os.path.join(self.my_path,'track_contacts.xlsx'),
                            sheetname = 'tracks', header = 0) if not debug else \
                read_excel(os.path.join(my_path,'track_contacts.xlsx'), sheetname = 'try1', header = 0)
        self.tracks = read_pickle(os.path.join(self.my_path,'shecodes_academy27.dat'))
        try:
            self.attendance = read_pickle(os.path.join(self.my_path,'LessonState27.dat')) if not debug else \
                                self.attdebug()
        except (IOError, ValueError):
            self.attendance = read_excel(os.path.join(self.my_path,'LessonState.xlsx'),
                                         sheetname = 'Sheet1', header = 0)
            self.attendance.to_pickle(os.path.join(self.my_path,'LessonState27.dat'))
            print('LessonState file was probably changed. Changes were saved.')
            print('')
        self.today = datetime.now().strftime('%d/%m/%Y') 
        try: 
            self.attlist = read_pickle(os.path.join(self.my_path,'attlist27.dat'))
            if self.today not in self.attlist.columns:
                if not members_update:
                    self.attlist[self.today] = np.nan
        except IOError:
            self.attlist = DataFrame(columns=['Name', self.today,
                                              'Studying','Starting month', 'Lesson'])
            self.attlist = self.attlist.set_index('Name')
        try:
            self.former = read_excel(os.path.join(self.my_path, 'FormerMembersInf.xlsx'))
        except IOError:
            pass
        try:
            self.insp_mem = read_pickle(os.path.join(self.my_path,'inspected_mem27.dat'))
            self.insp_mem = list(self.insp_mem)
        except IOError:
            self.insp_mem = []
    def attdebug(self):
        #======================
        # Just debug function
        #======================
        self.attendance = DataFrame(columns=[i for i in range(1,33)], 
                                    data= [['nan']*len(range(1,33))]*len(self.w.FullName.values),
                                    index = self.w.FullName.values)
        for name in self.w.FullName.values:
            reg_day = self.w[self.w.FullName==name].Timestamp[self.w[self.w.FullName==name].Timestamp.index[0]]
            self.attendance.loc[name,1] = str(reg_day.month)+'/'+str(reg_day.day)+'/'+str(reg_day.year)
        return self.attendance
    
    def check_if_under_inspection(self, name):
        self.name = name.title()
        if self.name in self.insp_mem:
            print(self.error)
            while 1:
                print('')
                print("I'm sorry :(")
                print('An error occurred during identification process of '+self.name+'.')
                print('')
                p1 = getpass.getpass('Insert password: ')
                if p1=='shecodes':
                    p2 = getpass.getpass("Enter manager's password: ")
                    if p2=='knvzvktguavkh.com':
                        print('')
                        what_to_do = raw_input('EXIT ? (y/n) ')
                        if what_to_do =='y':
                            break
                        if what_to_do=='n':
                            self.add_attendance(name, verbose=True)
                            break
                    else:
                        print('')
                        print('***Error: Authorization has been denied for this action***')
                        print('')
                else:
                    print('')
                    print('***Error: Authorization has been denied for this action***')
                    print('')
        else:
            self.add_attendance(self.name, verbose=True) 
        
    
    def add_attendance(self, name, verbose=True):
        #=======================================================================
        # Identify member + add attendance + interactive sending (yes/no) tasks
        #=======================================================================
        self.name = name.title()
        if self.name in self.attendance.index:
            self.short = re.search(r'\S+', self.name)
            if verbose: 
                print('\n Hi '+ self.short.group()+ ' :) \n')
            else: 
                print('You are logged to '+self.short.group()+ "'s account \n")

            self.to_send = raw_input('Do you want to get the next tasks? (Yes/No)')
            #===============================================================
            #              Sending the next tasks
            #===============================================================
            #Read self.send_req_lesson(verbose) that create+send tasks email
            if (self.to_send  == 'yes') or (self.to_send=='y') or (self.to_send=='Y') or (self.to_send=='YES') or (self.to_send=='Yes'):
                self.to_send = 'y'
                self.lesson = 0
                #---------------------------------------------
                # I cancel the first lesson option for now...
                #---------------------------------------------
#                 if (self.attendance.loc[self.name][2]=='nan') or \
#                     (isnull(self.attendance.loc[self.name][2])):
#                     self.lesson = self.first_lesson(verbose)
                if (self.lesson == 2) or (self.lesson==0):
                #TODO: cancel the above condition if you cancel "first lesson" option
                    for i in self.attendance.loc[self.name].index:
                        if (self.attendance.loc[self.name][i]=='nan') or\
                            (isnull(self.attendance.loc[self.name][i])):
                            self.lesson = i
                            break
                now = datetime.now()
                self.attendance.loc[self.name, self.lesson] = now.strftime('%d/%m/%Y') if verbose \
                                                                else 'missed'
                if self.lesson >1:
                    if self.attendance.loc[self.name, self.lesson] == self.attendance.loc[self.name, self.lesson-1]:
                        self.lesson = self.authorization_block()
                        print("Sending tasks for meetup number "+str(self.lesson))
                self.send_req_lesson(verbose) 
                self.attendance.to_pickle(os.path.join(self.my_path, 'LessonState.dat'))
            #=====================================================================
            # Don't want the next tasks - sending an attendance confirmation email
            #=====================================================================
            if (self.to_send  == 'no') or (self.to_send=='n') or (self.to_send=='N') or (self.to_send=='NO') or (self.to_send=='No'):
                self.to_send = 'n'
                if verbose:
                    print('Please wait...\n')
                    msgRoot = MIMEMultipart('related')
                    msgAlternative = MIMEMultipart()
                    msgRoot.attach(msgAlternative)
                    to = 'Hi '+self.short.group()+ ',\n'
                    ww = '\n'
                    body1 = 'You have successfully confirmed your attendance on '+self.today+'.\n'
                    body2 = "You haven't finished your previous tasks yet but it's okay.\n"
                    body3 = 'I just want you to know that if you finish your previous tasks today you can log in again and ask for the next tasks.\n'
                    prepre_end = "\n---------------------------------------------------IMPORTANT--------------------------------------------------- \n" 
                    pre_end = "\n*** Sign in on She Codes even if you didn't finished your tasks; ---> http://Bit.ly/shecodesacademy \n"
                    body = to+ww+body1+body2+body3+prepre_end+pre_end
                    msgText1 = MIMEText(body, 'plain')
                    msgText2 = MIMEText('<br><img src="cid:image1"><br><b>-------------------------<br><i>She codes(WIS)</i></b>', 'html')
                    msgAlternative.attach(msgText1)
                    msgAlternative.attach(msgText2)
                    msgRoot['Subject']= 'Shecodes; WIS Attendance Confirmation'
                    server = smtplib.SMTP('smtp.gmail.com', port = 587, timeout = 120)#587 465
                    server.ehlo()
                    server.starttls()
                    server.login("dar.shecodes", "shecodes")
                    server.sendmail("dar.shecodes@gmail.com", 
                                    self.w[self.w.FullName==self.name].email[self.w[self.w.FullName == self.name].index[0]], 
                                    msgRoot.as_string())  
                    server.quit()
                    print('OK. A confirmation email has been sent. Check your inbox and your Spam folder.\nLet us know if you have any problem.')
                    print('\nGood Luck :)\n')
            #===============================================================
            #              Add attendance in attlist.dat table
            #===============================================================
            #Note: You can't add attendance if verbose = false (= missing)
            #so if you send her the next tasks it won't recoded as attendance
            self.meetup = 0
            if verbose:
                for i in self.attendance.loc[self.name].index:
                        if (self.attendance.loc[self.name][i]=='nan') or\
                            (isnull(self.attendance.loc[self.name][i])):
                            self.meetup = i-1
                            break
                if self.name not in self.attlist.index:
                    s = Series({self.today: 'V',
                                'Studying': self.w[self.w.FullName == self.name].track[self.w[self.w.FullName == self.name].index[0]], 
                                'Starting month': self.w[self.w.FullName == self.name].Start[self.w[self.w.FullName == self.name].index[0]].strftime('%b, %Y'),
                                'Lesson': self.meetup})
                    s.name = self.name
                    self.attlist = self.attlist.append(s)
                elif self.name in self.attlist.index:
                    self.attlist.loc[self.name, self.today] = 'V'
                    self.attlist.loc[self.name, 'Lesson'] = self.meetup
#                 self.attlist = self.attlist.append({'Name':self.name,
#                                                     self.today: 'V',
#                                                     'Studying': self.w[self.w.FullName == self.name].track[self.w[self.w.FullName == self.name].index[0]],
#                                                     'Starting month': self.w[self.w.FullName == self.name].Start[self.w[self.w.FullName == self.name].index[0]].strftime('%b, %Y'), 
#                                                     'Lesson': self.meetup}, ignore_index=True)
                self.attlist.to_pickle(os.path.join(self.my_path, 'attlist.dat'))
        #==================================
        #       Don't know her name
        #==================================
        if (self.name not in self.attendance.index) and ((self.name!='quit') or (self.name!='q')):
            print('\nWe are sorry... :( \n\nYour tasks cannot be sent because we cannot find your name in our system.')
            print('Make sure you entered your name correctly and then try again.')
            print('\nIf you typed the name correctly, this error could be the result of:')
            print('(1) You are not yet registered to study track. If so:')
            print('    a. Register in our online form. Write your full name in English')
            print('    b. Talk with Dar or FooBar so they send you your tasks.')
            print('(2) You registered today but our system is out-to-date. Talk with Dar or Toot so they send you your tasks.')
            print('(3) When you filled out our registration form, you accidently selected another she codes branch instead of WIS.')
            print('    Talk with Dar or Toot to check this possibility.')
            print('\n****IMPORTANT****\nIf none of the above work for you, talk with Dar or Toot')
                   
    def first_lesson(self, verbose):
        # for now it canceled
        if verbose: print("I don't know if today is your first lesson, so please answer:" )
        first = 'yes'
        while 1:
            first = raw_input('Answer yes/no: Is this your first lesson? ') if verbose \
            else raw_input("Answer yes/no Is it her first lesson? " )
            if (first  == 'yes') or (first=='y') or (first=='Y') or (first=='YES') or (first=='Yes') or \
                        (first == 'no') or (first=='n') or (first=='N') or (first=='NO') or (first=='No'): 
                break
        if (first  == 'yes') or (first=='y') or (first=='Y') or (first=='YES') or (first=='Yes'):
            if verbose: print("\nSo... Welcome! Wishing you the very best of luck with your studies! <3 \n")
            else: print("Sending her the first lesson...\n")
            self.lesson = 1
        if (first == 'no') or (first=='n') or (first=='N') or (first=='NO') or (first=='No'):
            self.lesson = 2
        return self.lesson

    def send_req_lesson(self,verbose):
        #=============================
        # Creates + sends tasks email
        #=============================
        print('Please wait...\n')
        her_track = self.w[self.w.FullName == self.name].track[self.w[self.w.FullName == self.name].index[0]]
        if her_track != 'web_new' and her_track !='android_new':
            build_progress_bar(her_track, self.lesson)
        msgRoot = MIMEMultipart('related')
        msgAlternative = MIMEMultipart()
        msgRoot.attach(msgAlternative)
        ww = '\n'
        prepre_end = "\n------------------------------IMPORTANT------------------------------ \n" 
        pre_end = "\n*** Dont forget to sign in on She Codes; ---> http://Bit.ly/shecodesacademy \n"
        to = 'Hi '+self.short.group()+ ',\n' 
        if her_track == 'android_new':
            try:
                pdf = open(os.path.join(self.my_path,"new_android/"+str(self.lesson)+".pdf"), "rb").read()
                msgPdf = MIMEApplication(pdf, 'pdf')
                msgPdf.add_header('Content-Disposition', 'attachment', filename='Your tasks')
                msgRoot.attach(msgPdf)
                tasks = 'See the attached file (:'
            except IOError:
                tasks = self.tracks.ix[her_track, self.lesson]
        elif her_track == 'web_new':
            try:
                pdf = open(os.path.join(self.my_path,"new_web/"+str(self.lesson)+".pdf"), "rb").read()
                msgPdf = MIMEApplication(pdf, 'pdf')
                msgPdf.add_header('Content-Disposition', 'attachment', filename='Your tasks')
                msgRoot.attach(msgPdf)
                tasks = 'See the attached file (:'
            except IOError:
                tasks = self.tracks.ix[her_track, self.lesson]
        else:
            tasks = self.tracks.ix[her_track, self.lesson]
        body = to+ww+tasks+prepre_end+pre_end        
        msgText1 = MIMEText(body, 'plain')
        msgText2 = MIMEText('<br><img src="cid:image1"><br><b>-------------------------<br><i>She codes(WIS)</i></b>', 'html')
        msgAlternative.attach(msgText1)
        msgAlternative.attach(msgText2)
        if her_track != 'android_new' and her_track != 'web_new':
            with open('progressbar.png', 'rb') as fp:
                msgImage = MIMEImage(fp.read())
                fp.close()
        else:
            with open('progressbar_new.png', 'rb') as fp:
                msgImage = MIMEImage(fp.read())
                fp.close()
        msgImage.add_header('Content-ID', '<image1>')
        msgRoot.attach(msgImage)
        t = 'Android' if her_track=='android' else 'Android (Version 12/15)' if her_track=='android_new' else \
            'Web' if her_track=='web' else 'Web (Version 12/15)' if her_track=='web_new' else 'Short Web track'
        msgRoot['Subject']= 'Your tasks for meetup number '+str(self.lesson)+' : '+ t
        server = smtplib.SMTP('smtp.gmail.com', port = 587, timeout = 120)#587 465
        server.ehlo()
        server.starttls()
        server.login("my.email", "mypassword")
        server.sendmail("my.email@gmail.com", 
                        self.w[self.w.FullName==self.name].email[self.w[self.w.FullName == self.name].index[0]], 
                        msgRoot.as_string())  
        server.quit()
        if verbose: 
            print('OK. Your tasks have been sent. Check your inbox and your Spam folder.\nLet us know if you have any problem.')
            print('\nGood Luck :)\n')
        else:
            print('\nOK. Her tasks have been sent. Tell her to check her inbox and or Spam folder.\n')
        
    def authorization_block(self):
        print('You are about to send an advanced lesson, which requires special authorization\n')
        password = getpass.getpass("Enter manager's password: ")
        if password=='shecodes':
            password2 = getpass.getpass("Enter manager's password: ")
            if password2 =='knvzvktguavkh.com':
                return self.lesson
            else:
                print('\n***Error: Authorization has been denied for this request***\n')
                print('Talk with Dar in order to do that')
                self.attendance.loc[self.name, self.lesson]=np.nan
                self.lesson -=1
        else:
            print('***Error: Authorization Denied***')
            print('Talk with Dar in order to to that')
            self.attendance.loc[self.name, self.lesson]=np.nan
            self.lesson -=1
        return self.lesson
    
    def build_attend(self):
        df = DataFrame()
        try:
            df = read_pickle(os.path.join(self.my_path,'LessonState.dat'))
        except IOError:
            print('Creating attendance.dat file...')
        if df.empty:
            df = DataFrame(columns=[i for i in range(1,33)], 
                           data= [['nan']*len(range(1,33))]*len(self.w.FullName.values),
                           index = self.w.FullName.values)
            for name in self.w.FullName.values:
                reg_day = self.w[self.w.FullName==name].Timestamp[self.w[self.w.FullName==name].Timestamp.index[0]]
                # I have a bug with converting Timestamp in excel to python format
                # that switching between the day and month
                df.loc[name,1] = str(reg_day.month)+'/'+str(reg_day.day)+'/'+str(reg_day.year)
            df.to_pickle(os.path.join(self.my_path,'LessonState.dat'))
        else:
            names_to_add = set(self.w.FullName.values) - set(df.index)&set(self.w.FullName.values)
            names_to_add = set(names_to_add)-set(self.former.index)
            names_to_drop = set(df.index) - (set(df.index)&set(self.w.FullName))
            if len(names_to_add)>0:
                for name in names_to_add:
                    df.loc[name,:] = [np.nan]*32
                    reg_day = self.w[self.w.FullName==name].Timestamp[self.w[self.w.FullName==name].Timestamp.index[0]]
                    # I have a bug with converting Timestamp in excel to python format
                    # that switching between the day and month
                    df.loc[name,1] = str(reg_day.month)+'/'+str(reg_day.day)+'/'+str(reg_day.year)
                    if len(str(reg_day.month))==1 and len(str(reg_day.day))==1:
                        first_day = '0'+str(reg_day.month)+'/0'+str(reg_day.day)+'/'+str(reg_day.year) 
                    elif len(str(reg_day.month))==2 and len(str(reg_day.day))==1:
                        first_day = str(reg_day.month)+'/0'+str(reg_day.day)+'/'+str(reg_day.year)
                    elif len(str(reg_day.month))==1 and len(str(reg_day.day))==2:  
                        first_day = '0'+str(reg_day.month)+str(reg_day.day)+'/'+str(reg_day.year)
                    else:
                        first_day = str(reg_day.month)+'/'+str(reg_day.day)+'/'+str(reg_day.year)
                    self.attlist.loc[name, first_day]='V'
                    self.attlist.loc[name, 'Studying']= self.w[self.w.FullName == name].track[self.w[self.w.FullName == name].index[0]]
                    self.attlist.loc[name, 'Starting month']= self.w[self.w.FullName == name].Start[self.w[self.w.FullName == name].index[0]].strftime('%b, %Y')
                    self.attlist.loc[name, 'Lesson']= 1
                    print(name+' has been added')
            if len(names_to_drop)>0:
                df = df.drop(names_to_drop, axis=0)
                try:
                    self.attlist = self.attlist.drop(names_to_drop, axis=0)
                    for name in names_to_drop:
                        print(name+' has been deleted')
                except:
                    for name in names_to_drop:
                        try:
                            self.attlist = self.attlist.drop(name, axis=0)
                            print(name+' has been deleted')
                        except:
                            pass
            df.to_pickle(os.path.join(self.my_path, 'LessonState.dat'))
            self.attlist.to_pickle(os.path.join(self.my_path, 'attlist.dat'))
    
    def add_event(self):
        day_event = raw_input('Enter the day of an event (format dd/mm/yyyy): ')
        the_event = raw_input('The event (i.e Purim eve, Lecture at Hatahana Pub): ')
        self.attlist[day_event] = the_event
        self.attlist.to_pickle(os.path.join(self.my_path, 'attlist.dat'))
        print('attlist.dat was saved. Move it to your Drive.')
    
    def inspection_members(self):
        try:
            insp_members = read_pickle(os.path.join(self.my_path,'inspected_mem.dat'))
            if not insp_members.empty:
                print('')
                print('')
                print('Previously, members under inspection were: ')
                for member in insp_members.values:
                    print member
                print('')
                print('In this procedure the previous file will be deleted.')
                print('If you still want to inspect the attendance of the members above')
                print('add them to the new list.\n')
                print('Initializing list...\n')
            else:
                print('No one is under inspection\n')
        except:
            print('')
            print('No one is under inspection')
            print('')
        while 1:
            inspected = raw_input('The number of members under inspection: ')
            if inspected =='q':
                break
            else:
                try:
                    inspected = int(inspected)
                    insp_members = []
                    active_members = [n for n in self.attendance.index if n not in self.former.index]
                    dic = {i:name for (i, name) in enumerate(active_members)}
                    n_cols = math.ceil(len(dic.keys())/10.)
                    for c in range(int(n_cols)):
                        if c==0:
                            df = DataFrame({'col_0': dic.keys()[0:10], 'names_0':dic.values()[0:10]})
                        if c>0 and c<n_cols-1:
                            df['col_'+str(c)] = dic.keys()[10*(c):10+(10*(c))]
                            df['names_'+str(c)] = dic.values()[10*(c):10+(10*(c))]
                        if c==n_cols-1:
                            df['col_'+str(c)] = dic.keys()[10*(c):]+[np.nan]*(10-len(dic.keys()[10*(c):]))
                            df['names_'+str(c)] = dic.values()[10*(c):]+[np.nan]*(10-len(dic.values()[10*(c):]))
                    dfs_list = []
                    if len(df.columns)>8:
                        dfs_num = int(math.ceil(len(df.columns)/8.))
                        for i in range(dfs_num):
                            if i==0:
                                dfs_list.append(df[df.columns[0:8]])
                            else:
                                dfs_list.append(df[df.columns[8*i:8*(i+1)]])
                    else:
                        dfs_list.append(df)
                    for ddf in dfs_list:
                        rows_df = range(len(ddf))
                        cols_df = list(ddf)
                        sub_df = ddf.ix[rows_df,cols_df]
                        sub_dstr = sub_df.to_string(index=False, col_space=13).split('\n')
                        for line in sub_dstr[1:]:
                            print(line)
                        print('\n')
                    for i in range(inspected):
                        index_mem = raw_input('Print member '+str(i+1)+' by index: ')
                        if index_mem=='q':
                            break
                        else:
                            try:
                                index_mem = int(index_mem)
                                insp_members.append(dic[index_mem])
                                print('The attendance of '+dic[index_mem]+' is now under inspection.')
                            except:
                                stop = raw_input("Press 'f' to save if you finish or 'q' to delete and quit: ")
                                if stop=='f':
                                    break
                                elif stop=='q':
                                    insp_members = []
                                    break
                    insp_members = Series(insp_members)
                    insp_members.to_pickle(os.path.join(self.my_path,'inspected_mem.dat'))
                    break
                except:
                    print("Type a number to continue or 'q' to exit")
#########################################################################################################################
#                                   Confirm Registration
#########################################################################################################################                 
class ConfirmRegistration():
    def __init__(self, error_, my_path):
        self.error = error_
        self.my_path = my_path
        self.w = read_excel(os.path.join(self.my_path,'track_contacts.xlsx'),
                            sheetname = 'tracks', header = 0)
        try:
            self.attendance = read_pickle(os.path.join(self.my_path,'LessonState.dat'))
        except (IOError, ValueError):
            self.attendance = read_excel(os.path.join(self.my_path,'LessonState.xlsx'),
                                         sheetname = 'Sheet1', header = 0)
            self.attendance.to_pickle(os.path.join(self.my_path,'LessonState.dat'))
            print('LessonState file was probably changed. Changes were saved.')
            print('')
        self.today = datetime.now()
        self.wed = self.today - timedelta(days=self.today.weekday())+timedelta(days=2, weeks=-1)
        self.wed = self.wed.strftime('%d/%m/%Y')
        try: 
            self.attlist = read_pickle(os.path.join(self.my_path,'attlist.dat'))
        except IOError:
            self.attlist = DataFrame(columns=['Name', self.wed,
                                              'Studying','Starting month', 'Lesson'])
            self.attlist = self.attlist.set_index('Name')
        
    def send_confirmation_message(self, name):
        self.name = name.title()
        if self.name in self.attendance.index:
            self.short = re.search(r'\S+', self.name)
            print('\n Hi '+ self.short.group()+ ' :) \n')
            studies_opt = ['web', 'web_new', 'android', 'android_new']
            her_studies = self.w[self.w.FullName == self.name].track[self.w[self.w.FullName == self.name].index[0]]
            if her_studies not in studies_opt:
                print(self.error)
                print('')
                print('You are NOT registered to any course here!')
                print('Please register and try again.')
            else:
                self.studing = 'Android (V.12.15)' if her_studies == 'android_new' else 'Web (V.12.15)' if \
                            her_studies == 'web_new' else 'Web (V.8.15)' if her_studies == 'web' else 'Android (V.8.15)'
                self.class_ = self.w[self.w.FullName == self.name].Start[self.w[self.w.FullName == self.name].index[0]].strftime('%B %Y')
                print('Please wait...\n')
                msgRoot = MIMEMultipart('related')
                msgAlternative = MIMEMultipart()
                msgRoot.attach(msgAlternative)
                to = 'Hi '+self.short.group()+ ',\n'
                ww = '\n'
                body1 ="Welcome to She codes; WIS branch!"
                body1 = 'You have successfully registered to '+self.studing+' development studies: Class of '+self.class_+'.\n'
                body2 = 'Also, your attendance on '+self.wed+' is documented.\n'
                body3 = 'Good Luck :)\n'
                prepre_end = "\n---------------------------------------------------IMPORTANT--------------------------------------------------- \n" 
                pre_end = "\n*** Sign up to the international She Codes Academy; ---> http://Bit.ly/shecodesacademy \n"
                body = to+ww+body1+body2+body3+prepre_end+pre_end
                msgText1 = MIMEText(body, 'plain')
                msgText2 = MIMEText('<br><img src="cid:image1"><br><b>-------------------------<br><i>She codes(WIS)</i></b>', 'html')
                msgAlternative.attach(msgText1)
                msgAlternative.attach(msgText2)
                msgRoot['Subject']= '>>>Shecodes; WIS<<< Registration Confirmation'
                server = smtplib.SMTP('smtp.gmail.com', port = 587, timeout = 120)#587 465
                server.ehlo()
                server.starttls()
                server.login("dar.shecodes", "shecodes")
                server.sendmail("dar.shecodes@gmail.com", 
                                self.w[self.w.FullName==self.name].email[self.w[self.w.FullName == self.name].index[0]], 
                                msgRoot.as_string())  
                server.quit()
                print('OK. A confirmation email has been sent. Check your inbox and your Spam folder.\nLet us know if you have any problem.')
                print('\nGood Luck :)\n')
                raw_input('Press ENTER to exit')
                print('')
        else:
            print('')
            print(self.error)
            print("")
            print("Name Error: Can't find your name.")
            print("")
            print("Possible solutions:")
            print("-------------------")
            print("1. Have you spelled your name correctly?")
            print("   If you have move to the next section. If you havn't press ENTER and try again.")
            print("2. Act as follow:")
            print("   (a) Rewrite your name in the Excel table you registered and capitalize the")
            print("       first letter of your first name and your last name (e.g. Dar Lador).")
            print("       *** Important note: DO NOT use hyphen (i.e. - ) or dot or any other sign.")
            print("   (b) Talk with Dar to restart this program.")
            print("")
            print("")
            raw_input('Press ENTER to continue')
            print('')
            
#########################################################################################################################
#                                   Main
#########################################################################################################################       
       
def main():
    my_path = 'C:/Users/Dar1/Documents/she_codes/AcademyInfo'
    print(shecodes)
    while 1:
        p = getpass.getpass('Password: ')
        if (p=='quit') or (p=='q'):
            print('Bye Bye...')
            break 
        elif p=='1234' or (p=='debug'):
            p2 = getpass.getpass('Password: ')
            if p2=='1234' or (p2=='debug'):
                while 1:
                    print('-------------------------------------------------------------------------------------------------------------------------')
                    print("                                          *~*~*~*~*~*~*I'm the TasksMonster!*~*~*~*~*~*~*")
                    print('                                          (c) all rights reserved to Shecodes; WIS branch')
                    print('-------------------------------------------------------------------------------------------------------------------------')
                    print('')
                    if (p2=='debug') or (p=='debug'):
                        print('DEBUG Version is True!!!\n')
                    name = raw_input('Enter your full name: ')
                    if (p2=='debug') or (p=='debug'):
                        try:
                            Attend(error_, my_path, debug=True).check_if_under_inspection(name.lower())
                            #Attend(error_, my_path, debug=True).add_attendance(name.lower(), verbose=True)
                        except AttributeError:
                            print("\n***Error*** You must create an attendance.dat file!")
                            print("Type 'quit' on your keyboard, type the manager's passwords and choose option 'b'.")
                    elif (name == 'quit') or (name == 'q'):
                        print('Bye Bye...')
                        break
                    else:
                        try:
                            Attend(error_, my_path, debug=False).check_if_under_inspection(name.lower())
                            #Attend(error_, my_path, debug=False).add_attendance(name.lower(), verbose=True)
                        except AttributeError:
                            print("\n***Error*** You must create an attendance.dat file!")
                            print("Type 'quit' on your keyboard, type the manager's passwords and choose option 'b'.")
#                     if (name == 'quit') or (name == 'q'): 
#                         print('Bye Bye...')
#                         break
        elif p=='shecodes':
            p2 = getpass.getpass("Enter manager's password: ")
            if (p2 =='knvzvktguavkh.com') or (p2=='debug'):
                print('\nHey shecodes manager :)')
                if p2=='debug': print('DEBUG Version is True!!!\n')
                while 1:
                    print('\nManager options\n===============')
                    print("a --> Sending weekly tasks to programmers who missed the HackNigh, while setting 'missed' in the attendance.dat file.")
                    print('b --> Edit/Update or create the attendance.dat file for new registers shown in track_contacts.xlsx')
                    print('c --> Add an event')
                    print("d --> Inspection of members' login")
                    print("e --> Registration confirmation email")
                    print('info --> Information about the required structure of the excel track_contacts.xlsx table')
                    print('q --> Quit')
                    decision = raw_input('What would you like to do? ')
                    if decision=='a':
                        name = raw_input("Enter the Programmer's full name: ")
                        try:
                            if (p2=='debug'):
                                Attend(error_, my_path, debug=True).add_attendance(name.lower(), verbose=False)
                            else:
                                Attend(error_, my_path, debug=False).add_attendance(name.lower(), verbose=False)
                        except AttributeError:
                            print("\n***Error*** You must create an attendance.dat file!")
                            print("Type the manager's passwords and choose option 'b'.")
                    if decision=='b':
                        print("")
                        if (p2=='debug'):
                            Attend(error_, my_path, debug=1, members_update=False).build_attend()
                        else:
                            Attend(error_, my_path, debug=None, members_update=True).build_attend()
                    if decision=='c':
                        print("")
                        Attend(error_, my_path, debug=None).add_event()
                    if decision=='d':
                        Attend(error_, my_path, debug=None).inspection_members()
                    if decision=='e':
                        while 1:
                            print('')
                            print('-------------------------------------------------------------------------------------------------------------------------')
                            print("                                          *~*~*~*~*~*~*I'm the TasksMonster!*~*~*~*~*~*~*")
                            print("                                                Welcome to Shecodes; Academy - WIS branch")
                            print('                                                      (c) all rights reserved')
                            print('-------------------------------------------------------------------------------------------------------------------------')
                            print('')
                            name = raw_input('Enter your full name: ')
                            if (name == 'quit') or (name == 'q'): 
                                print('')
                                print('Bye Bye...')
                                break
                            else:
                                ConfirmRegistration(error_, my_path).send_confirmation_message(name.lower())
                    if decision=='info': print(xl_format)
                    if decision=='q': break
            else:
                print('***Error: Authorization has been denied for this request***')
                print('')
            
        
    

if __name__ == "__main__":
    main()