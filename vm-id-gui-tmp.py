__author__ = 'Admin'

import wx
import MySQLdb
import os
import time
import datetime
import glob,string
import csv
import pdb
from xlsxwriter.workbook import Workbook

#pdb.set_trace()

class project(wx.Frame):
    def __init__(self, parent, id):
        wx.Frame.__init__(self, parent, id, 'VM ID MAP GUI', size=(300, 300))
        self.panel = wx.Panel(self)

        """wx.StaticText(panel, -1, "Starting the report generation process..", (10,5))
        custom=wx.StaticText(panel, -1, "Below Showing is the hypervisors list:", (10,20))"""
        db_array = []
        global excel_count
        excel_count = 0
        excel_special = r'*:?/\[]'
        self.trans = string.maketrans(excel_special, ' '*len(excel_special))

        #Getting all the ESXi hosts on the GUI window
        mdb = MySQLdb.connect("10.163.221.161", "root", "0ps4w1n&l1n", "CRASH_REPORT_TABLES")
        cursor = mdb.cursor()
        cursor.execute("show tables")
        for items in cursor.fetchall():
            db_array.append(items[0])
        list(db_array)

        #Getting username from the user
        username_box=wx.TextEntryDialog(None, "Please enter the screen name", "Title", "Default")
        if username_box.ShowModal() == wx.ID_OK:
            answer = username_box.GetValue()
            #print answer

        #Setting the time stamp
        ts = time.time()
        st = datetime.datetime.fromtimestamp(ts).strftime('%Y_%m_%d_%H_%M_%S')
        #print str(st)
        self.session_stamp = str(st)+"_"+answer
        #print self.session_stamp
        self.dir_path = "/var/reports/" + str(self.session_stamp)
        os.makedirs(self.dir_path)
        os.chown(self.dir_path, 27, 27)

        #Chosing all the faulty ESXi from the list
        box = wx.MultiChoiceDialog(None, 'Here is a List of Hypervisors. Please make choices', 'VM-ORG-MAPPING', db_array)
        if box.ShowModal() == wx.ID_OK:
            ans = box.GetSelections()
            if not ans:
                self.Destroy()
            else:
                posy = 10
                posx = 10
                self.user_input = []
                for i in ans:
                    self.user_input.append(db_array[i])
                """posy = posy + 25
                #print db_array[i]
                wx.StaticText(panel, -1, 'You have made choices of the following hypervisors:', (10,10))
                wx.StaticText(panel, -1, str(db_array[i]), (posx, posy))"""
        else:
            pass
        
        box = wx.MessageDialog(None, 'Starting the process of report generation', 'Status',wx.OK)
        pop = box.ShowModal()
        box.Destroy()
        #CREATING TABLES IN DATABASE
    def creating_db_table(self):
        box1 = wx.MessageDialog(None, 'Creating DB tables:', 'Status',wx.OK)
        pop1 = box1.ShowModal()
        box1.Destroy()
        ss = self.session_stamp
        mdb = MySQLdb.connect("10.163.221.161","root","0ps4w1n&l1n","MAP_FINAL")
        cursor = mdb.cursor()
        #print "create table" ' %s' " (ORG_ID varchar(200), EMAIL_ADDRESS varchar(200),  FULL_NAME varchar(100),  DESCRIP varchar(200), VM_ID varchar(100))"%ss
        cursor.execute("create table" ' %s' " (ORG_ID varchar(200),EMAIL_ADDRESS varchar(200),FULL_NAME varchar(100),DESCRIP varchar(200), VM_ID varchar(100))"%ss)
        mdb.close()

    def connecting_10_162_0_100(self,ips, ids):
        self.ips = ips
        self.ids = ids
        db1 = MySQLdb.connect( "10.162.0.100","crash_report","h634GghU4rt3f3","wsapi" )
        cursor = db1.cursor()
        #cursor.execute("select concat_ws('""',t1.ORG_ID,',', t1.EMAIL_ADDRESS,',',t1.FULL_NAME,',',t2.DESCRIPTION) from wsapi.OEC_ACCOUNT as t1, wsapi.OEC_ORGANIZATION as t2 where t2.ID=%s and t1.ORG_ID=%s",  [ids, ids])
        cursor.execute("select t1.ORG_ID, t1.EMAIL_ADDRESS,t1.FULL_NAME,t2.DESCRIPTION from wsapi.OEC_ACCOUNT as t1, wsapi.OEC_ORGANIZATION     as t2 where t2.ID=%s and t1.ORG_ID=%s",  [ids, ids])
        for orgid in cursor.fetchall():
            b = list(orgid)
            b.append(ips)
            mapdb = MySQLdb.connect("10.163.221.161","root","0ps4w1n&l1n","MAP_FINAL",use_unicode=1,charset="utf8")
            cursor = mapdb.cursor()
            cursor.execute("insert into " '%s ' "( ORG_ID, EMAIL_ADDRESS, FULL_NAME, DESCRIP, VM_ID ) values ( " '"%s", "%s", "%s", "%s", "%s"' " )"%(self.session_stamp, b[0], b[1], unicode(b[2], errors='ignore'), unicode(b[3], errors='ignore'), b[4]))
        db1.close()

    def fetching_distinct_id(self):
        mdb = MySQLdb.connect("10.163.221.161","root","0ps4w1n&l1n","MAP_FINAL")
        cursor = mdb.cursor()
        cursor.execute("select distinct ORG_ID from " ' %s'%self.session_stamp)
        for distinct_org in cursor.fetchall():
            loc = self.dir_path+"/"+distinct_org[0]+".csv"
            cursor.execute("select distinct EMAIL_ADDRESS, FULL_NAME, DESCRIP, VM_ID from " '%s' " where ORG_ID=" '"%s"' " into outfile " '"%s"' " fields terminated by ',' lines terminated by '\\n'"%(self.session_stamp, distinct_org[0], loc))
            #cursor.execute("select EMAIL_ADDRESS, group_concat( distinct FULL_NAME separator  \"#\"), group_concat( distinct DESCRIP separator \"#\" ), group_concat( distinct VM_ID separator \"#\") from " '%s' " where ORG_ID=" '"%s"' " group by EMAIL_ADDRESS into outfile " '"%s"' " fields terminated by ',' lines terminated by '\\n'"%(session_stamp, distinct_org[0], loc))
            #select  EMAIL_ADDRESS, group_concat( distinct FULL_NAME),group_concat(distinct DESCRIP), group_concat( distinct VM_ID separator "#") from 2013_09_02_11_20_39_abhradeep_opsource where ORG_ID="e6fa08bf-0b08-4496-aaa6-3e2197226ec9" group by email_address

    def creating_excel(self):
        box6=wx.MessageDialog(None, 'Creating Excel worksheets and workbook.', 'Status',wx.OK)
        pop6=box6.ShowModal()
        box6.Destroy()
        workbook = Workbook("/var/reports/" + self.session_stamp + ".xlsx")
        for csvfile in glob.glob(os.path.join('/var/reports/',self.session_stamp,'*.csv')):
            sheet_name=self.get_org_name_from_orgid(os.path.basename(csvfile))
            worksheet = workbook.add_worksheet(sheet_name)
            with open(csvfile, 'rb') as f:
                reader = csv.reader(f)
                for r, row in enumerate(reader):
                    for c, col in enumerate(i.decode("utf-8", "replace") for i in row):
                        y = [ i.decode("utf-8", "replace") for i in row ]
                        worksheet.write(r, c, col)
        workbook.close()
        box7=wx.MessageDialog(None, 'Excel file has been created!!', 'Status',wx.OK)
        pop7=box7.ShowModal()
        box7.Destroy()

    def get_org_name_from_orgid(self, orgid_file):
        global excel_count
        orgdb = MySQLdb.connect("10.163.221.161","root","0ps4w1n&l1n","MAP_FINAL")
        cursor = orgdb.cursor()
        cursor.execute("select distinct DESCRIP from " '%s' " where ORG_ID=" '"%s"'%(self.session_stamp, os.path.splitext(orgid_file)[0]))
        for orgname in cursor.fetchall():
            excel_count = excel_count + 1
            return repr(str(excel_count)+"_"+orgname[0][0:26:1].translate(self.trans))

    def connecting_10_163_221_161(self, esxi):
        box3=wx.MessageDialog(None, 'Connecting to DB "CRASH_REPORT_TABLES"', 'Status',wx.OK)
        pop3=box3.ShowModal()
        box3.Destroy()
        db = MySQLdb.connect( "10.163.221.161","root","0ps4w1n&l1n","CRASH_REPORT_TABLES" )
        cursor = db.cursor()
        cursor.execute('select * from %s' % esxi)
        box4=wx.MessageDialog(None, 'Connecting 10.162.0.100 for ORG-ID', 'Status',wx.OK)
        pop4=box4.ShowModal()
        box4.Destroy()
        for items in cursor.fetchall():
            self.connecting_10_162_0_100(items[0], items[3])
        db.close()
        box5=wx.MessageDialog(None, 'All ORGIDs have been fetched', 'Status',wx.OK)
        pop5=box5.ShowModal()
        box5.Destroy()

    def send_email(self):
        pass


    def start_execution(self):
        box2=wx.MessageDialog(None, 'Starting Actual execution', 'Status',wx.OK)
        pop2=box2.ShowModal()
        box2.Destroy()
        for hypervisor in self.user_input:
            self.connecting_10_163_221_161(hypervisor)
        self.fetching_distinct_id()
        self.creating_excel()
        box_final=wx.MessageDialog(None, 'You can download the file from ftp site ftp://10.163.221.160/pub ..', 'Status',wx.OK)
        pop_final=box_final.ShowModal()
        box_final.Destroy()

if __name__ == '__main__':
    app=wx.PySimpleApp()
    frame=project(parent=None, id=-1)
    #frame.Show()
    frame.creating_db_table()
    frame.start_execution()
    #app.MainLoop()
    #frame.destroy()
