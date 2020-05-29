import sys
import time
import os
import  shutil
import xlwt

#Script creator :Tom'li , 2019/12/15
path=sys.argv[1]
xformat_file_path_list=[]
xformat_file_path_list_s=[]
Name_list=[]
signal='$$'
def get_file_path_list(rootdir,xformat_file_path_list,formats):#scan all the rtf files,and get paths
    lister = os.listdir(rootdir)
    for i in range(0,len(lister)):
        path=os.path.join(rootdir,lister[i])
        if os.path.isfile(path):
            if get_file_format(path)==formats:#get each txt files' path
            	xformat_file_path_list.append(path)
        else:
            get_file_path_list(path,xformat_file_path_list,formats)
    return xformat_file_path_list


def get_file_path_list_s(rootdir,xformat_file_path_list_s,formats,Name_list):#for fail count only 
    lister = os.listdir(rootdir)
    for i in range(0,len(lister)):
        path=os.path.join(rootdir,lister[i])
        if os.path.isfile(path):
            if get_file_format(path)==formats:#get each txt files' path
            	if get_file_name(path) in Name_list:
            		xformat_file_path_list.append(path)
        else:
            get_file_path_list_s(path,xformat_file_path_list_s,formats)
    return xformat_file_path_list_s


def get_file_format(path):#get the format of  file
    temp=path.split('.')
    return temp[-1]

def get_file_name(path):#get the format of  file
    temp=path.split('.')[0].split('/')
    return temp


def get_radar_tag(path):
	result=[]
	with open(path,'r') as f:
		content=f.read()
		temp=content.split('\n')
		for i in temp:
			#print(i.split('> ')[-1])
			result.append(i.split('> ')[0].split('/')[-1])
	return result

def get_name(path):
	result=[]
	with open(path,'r') as f:
		content=f.read()
		temp=content.split('\n')
		for i in temp:
			#print(i.split('> ')[-1])
			result.append(i.split('> ')[-1])
	return result

def write_to_csv(path,f):
	radar_list=get_radar_tag(path)
	name_list=get_name(path)
	
	if len(radar_list) == len(name_list):
		with open((path.split('/')[-1]).split('.')[0]+'_'+time_stamp+'.txt','w') as temp_f:
			for i in range(len(radar_list)):
				f.write('Medium'+signal+radar_list[i]+signal+(name_list[i])+'\n')
				temp_f.write('Medium'+signal+radar_list[i]+signal+(name_list[i])+'\n')

	else:
		#print(len(radar_list))
		#print(len(name_list))
		print("Status:Fatal ERROR!")
	return ((path.split('/')[-1]).split('.')[0]+'_'+time_stamp)


time_stamp=time.strftime("%Y-%m-%d-%H_%M_%S",time.localtime())#creat time stamp
path_list=get_file_path_list(path,xformat_file_path_list,'txt')#read all rtf file
print('Status:Find '+str(len(path_list))+' txt files')
with open((path.split('/')[-1]).split('.')[0]+'_RTF_'+time_stamp+'.txt','w') as f:
	f.write('Station'+signal+'Radar'+signal+'Name'+signal+'Failed Rate'+'\n')
	for i in range(len(path_list)):
		#print('Status:Processing NO.'+str(i+1)+'rtf file')
		Name_list.append(write_to_csv(path_list[i],f))
		#print(Name_list)
	#print("Status;Done!")
	#print("Visit"+os.getcwd()+'/'+(path.split('/')[-1]).split('.')[0]+'_RTF_'+time_stamp+'.txt'+"get result")
#remove repeat item
def remove_repeat_item(old_path,new_path):
	readPath=old_path
	writePath=new_path
	lines_seen=set()
	outfiile=open(writePath,'a+',encoding='utf-8')
	f=open(readPath,'r',encoding='utf-8')
	for line in f:
	    if line not in lines_seen:
	        outfiile.write(line)
	        lines_seen.add(line)
print('Status:Processing... wait...')
remove_repeat_item(os.getcwd()+'/'+(path.split('/')[-1]).split('.')[0]+'_RTF_'+time_stamp+'.txt',os.getcwd()+'/'+(path.split('/')[-1]).split('.')[0]+'_Del_Repeated_RTF_'+time_stamp+'.txt')
os.remove(os.getcwd()+'/'+(path.split('/')[-1]).split('.')[0]+'_RTF_'+time_stamp+'.txt')
def get_fail_count(path_list,core_path):
	result=[]
	with open(core_path,'r') as core_f:
			#print('open core path ok!')
			for core_item in core_f.read().split('\n'):
				#print('!!!@@@')
				#print(core_item)
				if core_item != '':
					#print(core_item.split(',')[0])
					if core_item.split(signal)[0] == 'Medium':
						#print('find Medium !')
						temp=0
						for item in path_list:
							with open(os.getcwd()+'/'+item+'.txt','r') as f:
								#print('open txt ok!')
								content=f.read()
								if core_item.split(signal)[2] in content:
									#print('find content !')
									temp=temp+1
						result.append(temp)
	return result

fail_counts=[]
fail_counts=get_fail_count(Name_list,os.getcwd()+'/'+(path.split('/')[-1]).split('.')[0]+'_Del_Repeated_RTF_'+time_stamp+'.txt')
#print(fail_counts)
#write f/t infomation to csv :


def write_F_T_Info_to_final_csv(core_path,fail_counts,path_list):#
	with open(core_path,'r') as f1:
		with open(os.getcwd()+'/'+'_FINAL_'+time_stamp+'.csv','w') as f2:
			f2.write('Station'+signal+'Radar'+signal+'Name'+signal+'Failed Rate'+'\n')
			tag=0

			for item in f1.read().split('\n'):
				#print('enter f1 ok')
				#print(item)
				if item.split(signal)[0] == 'Medium':
					#print('find Medium')
					f2.write(item.split('\n')[0]+signal+str(fail_counts[tag])+'F/'+str(len(path_list))+'T'+'\n')
					#print('write f2 ok ')
					tag=tag+1
write_F_T_Info_to_final_csv(os.getcwd()+'/'+(path.split('/')[-1]).split('.')[0]+'_Del_Repeated_RTF_'+time_stamp+'.txt',fail_counts,path_list)

for item in Name_list:
	os.remove(os.getcwd()+'/'+item+'.txt')
os.remove(os.getcwd()+'/'+(path.split('/')[-1]).split('.')[0]+'_Del_Repeated_RTF_'+time_stamp+'.txt')

fs = xlwt.Workbook()
sheet1 = fs.add_sheet('RESULT',cell_overwrite_ok=True)
with open(os.getcwd()+'/'+'_FINAL_'+time_stamp+'.csv','r')as fss:
	contents=fss.read()
	for item in range(len(contents.split('\n'))):
		for i in range(len(contents.split('\n')[item].split(signal))):
			sheet1.write(item,i,contents.split('\n')[item].split(signal)[i])
fs.save(os.getcwd()+'/'+'_FINAL_'+time_stamp+'.xls')
os.remove(os.getcwd()+'/'+'_FINAL_'+time_stamp+'.csv')
print('Done!')
print('Visit : '+os.getcwd()+'/'+'_FINAL_'+time_stamp+'.xls'+' get your result file !')




