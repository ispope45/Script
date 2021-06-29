import os
import glob
import shutil

OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'

PATH = "\\Ddnas\\인프라사업부_검수서류\\20200204_서울시교육청 망분리 및 고도화사업_2차(KT)"

f_list = glob.glob("D:\\Desktop\\excel\\*")