#!/usr/bin/python

import os,time,datetime,MySQLdb
import glob,string
import csv
import codecs

from xlsxwriter.workbook import Workbook

#VARIABLE DECLARATION SECTION
global excel_count
excel_count = 0
excel_special = r'*:?/\[]'
trans = string.maketrans(excel_special, ' '*len(excel_special))

#ENTER USER NAME
name = raw_input("Enter your name: ")
str(name)

#HYPERVISOR LIST INPUT FROM USER
user_input = ["jnbc4evm04p", "amsc4evm06p", "ashc4evm02p"]
#user_input = ["sydc4evm08p"]
ts = time.time()
st = datetime.datetime.fromtimestamp(ts).strftime('%Y_%m_%d_%H_%M_%S')
session_stamp = str(st)+"_"+name
dir_path = "/tmp/"+session_stamp
os.makedirs(dir_path)
os.chown(dir_path, 27, 27)

#CREATING TABLES IN DATABASE
mdb = MySQLdb.connect("10.163.221.161","root","0ps4w1n&l1n","MAP_FINAL")
cursor = mdb.cursor()
cursor.execute("create table %s (ORG_ID varchar(200), EMAIL_ADDRESS varchar(200),  FULL_NAME varchar(100),  DESCRIP varchar(200), VM_ID varchar(100))" % session_stamp)
mdb.close()

def connecting_10_162_0_100(ips, ids):
    db1 = MySQLdb.connect( "10.162.0.100","crash_report","h634GghU4rt3f3","wsapi" )
    cursor = db1.cursor()
    #cursor.execute("select concat_ws('""',t1.ORG_ID,',', t1.EMAIL_ADDRESS,',',t1.FULL_NAME,',',t2.DESCRIPTION) from wsapi.OEC_ACCOUNT as t1, wsapi.OEC_ORGANIZATION as t2 where t2.ID=%s and t1.ORG_ID=%s",  [ids, ids])
    cursor.execute("select t1.ORG_ID, t1.EMAIL_ADDRESS,t1.FULL_NAME,t2.DESCRIPTION from wsapi.OEC_ACCOUNT as t1, wsapi.OEC_ORGANIZATION     as t2 where t2.ID=%s and t1.ORG_ID=%s",  [ids, ids])
    
    for orgid in cursor.fetchall():
        b = list(orgid)
        b.append(ips)
        mdb = MySQLdb.connect("10.163.221.161","root","0ps4w1n&l1n","MAP_FINAL")
        cursor = mdb.cursor()
        cursor.execute("insert into " '%s ' "( ORG_ID, EMAIL_ADDRESS, FULL_NAME, DESCRIP, VM_ID ) values ( " '"%s", "%s", "%s", "%s", "%s"' " )"%(session_stamp, b[0], b[1], b[2], b[3], b[4]))
    db1.close()

def fetching_distinct_id():
    mdb = MySQLdb.connect("10.163.221.161","root","0ps4w1n&l1n","MAP_FINAL")
    cursor = mdb.cursor()
    cursor.execute("select distinct ORG_ID from " ' %s'%(session_stamp))
    for distinct_org in cursor.fetchall():
        loc = dir_path+"/"+distinct_org[0]+".csv"
        #print r"select EMAIL_ADDRESS, FULL_NAME, DESCRIP, VM_ID from " '%s' " where ORG_ID is " '"%s"' " into outfile " '"%s"' " fields terminated by ',' enclosed by '\"' lines terminated by '\\n'"%(table_name, distinct_org[0], loc)
        #cursor.execute("select distinct EMAIL_ADDRESS, FULL_NAME, DESCRIP, VM_ID from " '%s' " where ORG_ID=" '"%s"' " into outfile " '"%s"' " fields terminated by ',' lines terminated by '\\n'"%(session_stamp, distinct_org[0], loc))
        cursor.execute("select EMAIL_ADDRESS, group_concat( distinct FULL_NAME separator  \"#\"), group_concat( distinct DESCRIP separator \"#\" ), group_concat( distinct VM_ID separator \"#\") from " '%s' " where ORG_ID=" '"%s"' " group by EMAIL_ADDRESS into outfile " '"%s"' " fields terminated by ',' lines terminated by '\\n'"%(session_stamp, distinct_org[0], loc))
        #select  EMAIL_ADDRESS, group_concat( distinct FULL_NAME),group_concat(distinct DESCRIP), group_concat( distinct VM_ID separator "#") from 2013_09_02_11_20_39_abhradeep_opsource where ORG_ID="e6fa08bf-0b08-4496-aaa6-3e2197226ec9" group by email_address


def creating_excel():
    workbook = Workbook('/tmp/' + session_stamp + '.xlsx')
    for csvfile in glob.glob(os.path.join('/tmp',session_stamp,'*.csv')):
        sheet_name=get_org_name_from_orgid(os.path.basename(csvfile))
        worksheet = workbook.add_worksheet(sheet_name)
        with open(csvfile, 'rb') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, col in enumerate(i.decode("utf-8", "replace") for i in row):
                    y = [ i.decode("utf-8", "replace") for i in row ]
                    print col
                    worksheet.write(r, c, col)
    workbook.close()

def get_org_name_from_orgid(orgid_file):
    global excel_count
    orgdb = MySQLdb.connect("10.163.221.161","root","0ps4w1n&l1n","MAP_FINAL")
    cursor = orgdb.cursor()
    cursor.execute("select distinct DESCRIP from " '%s' " where ORG_ID=" '"%s"'%(session_stamp, os.path.splitext(orgid_file)[0]))
    for orgname in cursor.fetchall():
        excel_count = excel_count + 1
        return repr(str(excel_count)+"_"+orgname[0][0:26:1].translate(trans))


def connecting_10_163_221_161(esxi):
    db = MySQLdb.connect( "10.163.221.161","root","0ps4w1n&l1n","CRASH_REPORT_TABLES" )
    cursor = db.cursor()
    cursor.execute('select * from %s' % esxi)
    for items in cursor.fetchall():
         connecting_10_162_0_100(items[0], items[3])
    db.close()

for hypervisor in user_input:
    connecting_10_163_221_161(hypervisor)

fetching_distinct_id()
creating_excel()
