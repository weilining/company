import openpyxl
import datetime
import os
from openpyxl.utils import get_column_letter, column_index_from_string

user = "邓杰"
condition = "BGM3"
wafer_id = "T1AG87.21"
marking = "AD0018AU2-ES"
dut_id = 10
pm_pos_start_list    = {'1.56':'L20', '1.6':'L48', '1.64':'L83', '1.68':'L113', '1.72':'L144'}
nonce_pos_start_list = {'1.56':'H26', '1.6':'H59', '1.64':'H89', '1.68':'H119', '1.72':'H150'}

base_options = "--bmsc-options 115200:0 -o stratum.f2pool.com:3333 -u xuelei5151.1 -T --set-dc-com COM6:9600 --pattern "
special_options="--bmsc-rdreg 0000 --bmsc-vil 1 --bmsc-ctype 1390 --bmsc-blken 52:1 --bmsc-domain 0 --bmsc-domain-cross "
log_options="--logfile 1393.log --logfile-openflag w "

for volt in pm_pos_start_list.keys():
    extra_options = "--bmsc-setpll 1390:{freq} --cp2103-reset-chip --bmsc-voltage {volt_value} --bmsc-scan-freq-step {scan_freq_step} --bmsc-scan-freq-nonce-num {scan_freq_nonce_num} ".format(
                freq=300,
                volt_value=float(volt),
                scan_freq_nonce_num=208,
                scan_freq_step=50
    )
    run_cgminer_str="cgminer-1393.exe "+base_options+special_options+extra_options+log_options
    state_cur_name="state_cur_"+volt
    nonce_state_cur_name="nonce_state_cur_"+volt
    os.system(run_cgminer_str)
    os.system("cat 1393.log | grep ':golden nonce' | awk '{ printf $5\"\\t\"$9 \"\\n\"}' | awk -F, '{ printf $1\"\\n\" }' | awk '{nonce[$1]=nonce[$1](\"\\t\"$2);sum[$1]+=$2}END{for(i in nonce)print i,\"\\t\"sum[i],nonce[i]}' | sort > nonce_state")
    os.system('''cat 1393.log | grep 'After test pattern, Get DC current is'| awk '{printf "%.4f\\n",$10}' > '''+state_cur_name)
    os.system('''cat 1393.log | grep 'Before open core, Get DC current is'| awk '{printf "%.4f\\n",$10}' > leakage_cur''')
    os.system('''awk '{s=$1;getline<"nonce_state";print s"\\t",$0}' '''+state_cur_name+''' | awk '{s=$0;getline<"leakage_cur";print $1"\\t"s}' > '''+nonce_state_cur_name)

wb = openpyxl.load_workbook('temp.xlsx')
sheet = wb.copy_worksheet(wb.worksheets[0])

for volt in pm_pos_start_list.keys():
    fopen = open("pm_data_"+volt, 'r')
    pm_lines = fopen.readlines()
    fopen.close()
    fopen = open("nonce_state_cur_"+volt, 'r')
    nonce_state_lines = fopen.readlines()
    fopen.close()

    i=0
    for line in pm_lines:
        li=line.split('\t')
        val=int(li[0])
        sheet.cell(int(pm_pos_start_list[volt][1:]), column_index_from_string(pm_pos_start_list[volt][:1])+i, val)
        i=i+1

    i=0
    for line in nonce_state_lines:
        li=line.split('\t')
        j=0
        for l in li:
            val=float(l)
            sheet.cell(int(nonce_pos_start_list[volt][1:])+i, column_index_from_string(nonce_pos_start_list[volt][:1])+j, val)
            j=j+1
        i=i+1

sheet['B22'] = 0.4
sheet['C22'] = 0.8
sheet['D22'] = 1.2

wafer_info = wafer_id.split('.')
date_str = str(datetime.datetime.today())

sheet['B3'] = user
sheet['E3'] = date_str[:16]
sheet['A5'] = "1 芯片信息:("+condition+")"
sheet['K7'] = wafer_info[0]
sheet['L7'] = int(wafer_info[1])
sheet['B14'] = "("+str(dut_id)+"号板)"
sheet.title = wafer_id+"#"+str(dut_id)

wb.save("BM1393测试_"+condition+"_SLT_"+date_str[5:10]+".xlsx")
#wb.save('test.xlsx')


