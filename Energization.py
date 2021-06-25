
import psspy
import numpy 
import pandas as pd
import itertools
import time
import os,sys
import xlrd
from xlrd import open_workbook
import xlwt
import itertools
import pandas as pd
from openpyxl import load_workbook

_i=psspy.getdefaultint()
_f=psspy.getdefaultreal()
_s=psspy.getdefaultchar()


####################################################################################################
Scn = 3
V_POI = 1.0
# OLTC fixed ratio
oltc_ratio_ons = 0.93
oltc_ratio_ofs = 0.98
psspy.lines_per_page_one_device(1,1000000)
psspy.progress_output(2,r"""output_record.txt""",[0,0])
####################################################################################
#READ ARRAY INPUT DATA
####################################################################################
# excel file with collector array data
filename = "Collector_Sc1.xls"
All_data = pd.read_excel(filename,sheet_name = 'Scenario'+str(Scn))
String_num = All_data['String'].tolist()
WT_num = All_data['WTnum'].tolist()
WT_bus = All_data['WTBus'].tolist()
Node = All_data['ASOWNode'].tolist()
Cable_type = All_data['CableType'].tolist()
Length = All_data['Length'].tolist()
FromBus = All_data['FromBus'].tolist()
ToBus = All_data['ToBus'].tolist()
Off_num = All_data['Offnum'].tolist()
########################################################################################
####################################################################################
#PLANT BUS NUMBERS
####################################################################################

if Scn == 1 or Scn == 2 or Scn == 4:
	POI_bus = 227900 #CARDIFF  	
else: 
	POI_bus = 206294 #Larrabee


ONSS1_HV = 10101
ONSS2_HV = 10102
#ONSS3_HV = 10103
#ONSS4_HV = 10104
ONSS1_LV = 10201
ONSS2_LV = 10202
#ONSS3_LV = 10203
#ONSS4_LV = 10204

J_TUBE_2 = 10502
J_TUBE_1 = 10501
#OFSS3_HV = 10303
#OFSS4_HV = 10304
OFSS1_LV = 10401
OFSS2_LV = 10402
#OFSS3_LV = 10403
#OFSS4_LV = 10404

ONSRE_HDD_26 = 101202
ONSRE_HDD_16 = 101201

ONSRE_HDD_25 = 101102
ONSRE_HDD_15 = 101101

ONSRE_HDD_24 = 101002
ONSRE_HDD_14 = 101001

ONSRE_HDD_23 = 10902
ONSRE_HDD_13 = 10901

ONSRE_HDD_22 = 10802
ONSRE_HDD_12 = 10801

ONSRE_HDD_21 = 10702
ONSRE_HDD_11 = 10701

CNCRT_DB_2 = 90102
CNCRT_DB_1 = 90101

LNDFAL_HDD_2 = 10302
LNDFAL_HDD_1 = 10301

SEABED_2 = 10602
SEABED_1 = 10601

if Scn == 1 or Scn == 3:
	num_groups= 2
	ONSS_HV =[ONSS1_HV, ONSS2_HV]
	ONSS_LV =[ONSS1_LV, ONSS2_LV]
	OFSS_HV =[J_TUBE_1, J_TUBE_2]
	OFSS_LV =[OFSS1_LV, OFSS2_LV]
	ONSRE_HDD_6 =[ONSRE_HDD_16, ONSRE_HDD_26]
	ONSRE_HDD_5 =[ONSRE_HDD_15, ONSRE_HDD_25]
	ONSRE_HDD_4 =[ONSRE_HDD_14, ONSRE_HDD_24]
	ONSRE_HDD_3 =[ONSRE_HDD_13, ONSRE_HDD_23]
	ONSRE_HDD_2 =[ONSRE_HDD_12, ONSRE_HDD_22]
	ONSRE_HDD_1 =[ONSRE_HDD_11, ONSRE_HDD_21]
	CNCRT_DB =[CNCRT_DB_1, CNCRT_DB_2]
	LNDFAL_HDD =[LNDFAL_HDD_1, LNDFAL_HDD_2]
	SEABED =[SEABED_1, SEABED_2]

if Scn == 2 or Scn == 5:
	num_groups= 3
	ONSS_HV =[ONSS1_HV, ONSS2_HV, ONSS3_HV]
	ONSS_LV =[ONSS1_LV, ONSS2_LV, ONSS3_LV]
	OFSS_HV =[OFSS1_HV, OFSS2_HV, OFSS3_HV]
	OFSS_LV =[OFSS1_LV, OFSS2_LV, OFSS3_LV]
	SHORE = [SHORE1, SHORE2, SHORE3]
	HDD = [HDD1, HDD2, HDD3]
if Scn == 5:
	num_groups = 4
	ONSS_HV =[ONSS1_HV, ONSS2_HV, ONSS3_HV, ONSS4_HV]
	ONSS_LV =[ONSS1_LV, ONSS2_LV, ONSS3_LV, ONSS4_LV]
	OFSS_HV =[OFSS1_HV, OFSS2_HV, OFSS3_HV, OFSS4_HV]
	OFSS_LV =[OFSS1_LV, OFSS2_LV, OFSS3_LV, OFSS4_LV]
	SHORE = [SHORE1, SHORE2, SHORE3, SHORE4]
	HDD = [HDD1, HDD2, HDD3, HDD4]

Case_name = "Larrabee_MW_MVOW_1.0.sav"
Results_file = "Energization_Scenario"+str(Scn)+"VPOI"+str(V_POI)
psspy.case(Case_name)

####################################################################################################


################################ De-energize#######################################################

for i in range (len(WT_num)):
	# 
	psspy.dscn(ToBus[i])
	psspy.dscn(WT_bus[i])	
	psspy.fdns([0,0,0,1,1,0,0,0])

for i in range (num_groups):
	psspy.dscn(ONSS_HV[i])
	psspy.dscn(ONSS_LV[i])
	psspy.dscn(OFSS_HV[i])
	psspy.dscn(OFSS_LV[i])
	psspy.dscn(ONSRE_HDD_6[i])
	psspy.dscn(ONSRE_HDD_5[i])
	psspy.dscn(ONSRE_HDD_4[i])
	psspy.dscn(ONSRE_HDD_3[i])
	psspy.dscn(ONSRE_HDD_2[i])
	psspy.dscn(ONSRE_HDD_1[i])
	psspy.dscn(CNCRT_DB[i])
	psspy.dscn(LNDFAL_HDD[i])
	psspy.dscn(SEABED[i])
	psspy.machine_chng_2(ONSS_HV[i],r"""1""",[0,_i,_i,_i,_i,_i],[0.0,0.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""1""",0,[_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""2""",0,[_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""3""",0,[_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""4""",0,[_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""5""",0,[_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""6""",0,[_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""7""",0,[_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""8""",0,[_f,_f])
	psspy.shunt_chng(OFSS_HV[i],r"""1""",0,[_f,_f])
psspy.fnsl([0,0,0,1,1,0,0,0])
psspy.fnsl([0,0,0,1,1,0,0,0])

################################################################################################
#functions
################################################################################################
# Results dictionnary
result_dict = {}
def read_voltages(_buses):
	voltages = []
	for bus in _buses:
		ierr, type = psspy.busint(bus,'TYPE')
		if type < 4: 
			ierr, Vtemp = psspy.busdat(bus, 'PU')
		else:
			Vtemp = "None"
		voltages.append(Vtemp)
	return voltages

def max_voltages(_voltages):
	# # if all "None", return "None"
	if all(isinstance(v, str) for v in _voltages):
		max_voltages = "None"
	# if some "None" and some values, return max of values
	#if len(_voltages) > 0:
	else:
		max_voltages = round(max(v for v in _voltages if isinstance(v,float)),4)
	return max_voltages

def read_mach_output(bus, ID):
	ierr, type = psspy.busint(bus,'TYPE')
	if type < 4:
		ierr, Qgen_WT = psspy.macdat(bus,str(ID),'Q')
		ierr, Pgen_WT = psspy.macdat(bus,str(ID),'P')
	else:
		Qgen_WT = "None"
		Pgen_WT = "None"
	return Pgen_WT, Qgen_WT

def read_mach_Q(_bus):
	q_wt = []
	for mach in _bus:
		ptemp, qtemp = read_mach_output(mach,r"1")
		q_wt.append(qtemp)
	return q_wt

def max_q(_q):
	# if all "None", return "None"
	if all(isinstance(q, str) for q in _q):
		max_q = "None"
	# if some "None" and some values, return max of values
	else:
		max_q = round(max(q for q in _q if isinstance(q,float)),2)
	return max_q	

def get_results(_name):
	results1 = read_voltages(ONSS_HV)
	results2 = read_voltages(ONSS_LV)
	results3 = read_voltages(ONSRE_HDD_6)
	results4 = read_voltages(ONSRE_HDD_5)
	results5 = read_voltages(ONSRE_HDD_4)
	results6 = read_voltages(ONSRE_HDD_3)
	results7 = read_voltages(ONSRE_HDD_2)
	results8 = read_voltages(ONSRE_HDD_1)
	results9 = read_voltages(CNCRT_DB)
	results10 = read_voltages(LNDFAL_HDD)
	results11 = read_voltages(SEABED)
	results12 = read_voltages(OFSS_HV)
	results13 = read_voltages(OFSS_LV)
	res_66kV = read_voltages(ToBus)
	res_WT = read_voltages(WT_bus)
	Pstat, Qstat = read_mach_output(ONSS1_HV,r"1")
	q_wt = read_mach_Q(WT_bus)
	Converged = psspy.solved()
	results = [_name]
	results.append(max_voltages(results1))
	results.append(min(results1))
	results.append(max_voltages(results2))
	results.append(min(results2))
	results.append(max_voltages(results3))
	results.append(min(results3))
	results.append(max_voltages(results4))
	results.append(min(results4))
	results.append(max_voltages(results5))
	results.append(min(results5))
	results.append(max_voltages(results6))
	results.append(min(results6))
	results.append(max_voltages(results7))
	results.append(min(results7))
	results.append(max_voltages(results8))
	results.append(min(results8))
	results.append(max_voltages(results9))
	results.append(min(results9))
	results.append(max_voltages(results10))
	results.append(min(results10))
	results.append(max_voltages(results11))
	results.append(min(results11))
	results.append(max_voltages(results12))
	results.append(min(results12))
	results.append(max_voltages(results13))
	results.append(min(results13))
	results.append(max_voltages(res_66kV))
	results.append(min(res_66kV))
	results.append(max_voltages(res_WT))
	results.append(min(res_WT))	
	results.append(max_q(q_wt))
	results.append(min(q_wt))
	results.append(Qstat)
	results.append(Converged)
	result_dict[k]=results
################################ Energize#######################################################
psspy.plant_data(POI_bus,0,[V_POI, 100.0])

# Fix WT voltage at 1.00 pu and set at 0 MW, 0 MVAr
for i in range (len(WT_num)):
	psspy.plant_data(WT_bus[i],0,[1.00, 100.0])
	psspy.machine_chng_2(WT_bus[i],r"""1""",[1,_i,_i,_i,_i,_i],[0.0,0.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
	psspy.fdns([0,0,0,1,1,0,0,0])
	psspy.fdns([0,0,0,1,1,0,0,0])

# 1
# Reconnect onshore bus HV
for i in range (num_groups):
	psspy.recn(ONSS_HV[i])
	# Auto-adjustment of taps is disabled	
	psspy.two_winding_chng_6(ONSS_HV[i],ONSS_LV[i],r"""1""",[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,oltc_ratio_ons,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)
	psspy.fdns([0,0,0,1,1,0,0,0])
	psspy.fdns([0,0,0,1,1,0,0,0])

k=0
step_name = 'Conncect Onshore Bus HV'
get_results(step_name)

# 2
# Reconnect onshore bus LV
for i in range (num_groups):
	psspy.recn(ONSS_LV[i])
	psspy.fdns([0,0,0,1,1,0,0,0])
	psspy.fdns([0,0,0,1,1,0,0,0])
k+=1
step_name = 'Connect Onshore Bus LV'
get_results(step_name)

#loop for each branch
v_poi = read_voltages([POI_bus])
for i in range (num_groups):
	# 3
	# Reconnect STATCOM with vset = POI bus	
	psspy.plant_data(ONSS_HV[i],_i,[v_poi[0], 100.0])
	psspy.machine_chng_2(ONSS_HV[i],r"""1""",[1,_i,_i,_i,_i,_i],[0.0,0.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])	
	psspy.fdns([0,0,0,1,1,0,0,0])
	psspy.fdns([0,0,0,1,1,0,0,0])
	k+=1
	step_name = 'Connect STATCOM'+str(i+1)
	get_results(step_name)

	# 4
	# TURN ON 1 OR 2 SHUNTS (ADJUST DEPENDING ON CASE)	
	psspy.shunt_chng(ONSS_LV[i],r"""1""",1,[_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""2""",1,[_f,_f])
	#psspy.shunt_chng(ONSS_LV[i],r"""3""",1,[_f,_f])
	psspy.fdns([0,0,0,1,1,0,0,0])
	psspy.fdns([0,0,0,1,1,0,0,0])
	k+=1
	step_name = 'Connect Onshore Shunt 1 at Group'+str(i+1)
	get_results(step_name)

	# 5
	# Reconnect export cable
	psspy.recn(OFSS_HV[i])
	psspy.recn(ONSRE_HDD_6[i])
	psspy.recn(ONSRE_HDD_5[i])
	psspy.recn(ONSRE_HDD_4[i])
	psspy.recn(ONSRE_HDD_3[i])
	psspy.recn(ONSRE_HDD_2[i])
	psspy.recn(ONSRE_HDD_1[i])
	psspy.recn(CNCRT_DB[i])
	psspy.recn(LNDFAL_HDD[i])
	psspy.recn(SEABED[i])	
	psspy.fdns([0,0,0,1,1,0,0,0])
	psspy.fdns([0,0,0,1,1,0,0,0])
	k+=1
	step_name = 'Connect Export Cable at Group'+str(i+1)
	get_results(step_name)

	# 6
	# TURN ON OFFSHORE SHUNT
	psspy.shunt_chng(OFSS_HV[i],r"""1""",1,[_f,_f])	
	psspy.fdns([0,0,0,1,1,0,0,0])
	psspy.fdns([0,0,0,1,1,0,0,0])
	k+=1
	step_name = 'Connect Offshore Shunt at Group'+str(i+1)
	get_results(step_name)

	# 6a
	# TURN ON second onshore shunt
	psspy.shunt_chng(ONSS_LV[i],r"""3""",1,[_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""4""",1,[_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""5""",1,[_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""6""",1,[_f,_f])
	psspy.shunt_chng(ONSS_LV[i],r"""7""",1,[_f,_f])
	#psspy.shunt_chng(ONSS_LV[i],r"""8""",1,[_f,_f])
	psspy.fdns([0,0,0,1,1,0,0,0])
	psspy.fdns([0,0,0,1,1,0,0,0])
	k+=1
	step_name = 'Connect Onshore Shunt2, 3 and 4 at Group'+str(i+1)
	get_results(step_name)


	# 7
	# Reconnect offshore station lv
	psspy.recn(OFSS_LV[i])
	psspy.two_winding_chng_6(OFSS_HV[i],OFSS_LV[i],r"""1""",[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,oltc_ratio_ofs,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)
	psspy.fdns([0,0,0,1,1,0,0,0])
	psspy.fdns([0,0,0,1,1,0,0,0])
	k+=1
	step_name = 'Connect Offshore LV Bus at Group'+str(i+1)
	get_results(step_name)

	# 8.... Reconnect strings 
	for m in range (len(WT_num)):		
		if Off_num[m] == i+1:  #if string is connected to offshore bus i
			psspy.recn(ToBus[m])			
			psspy.fdns([0,0,0,1,1,0,0,0])
			psspy.fdns([0,0,0,1,1,0,0,0])
			psspy.fdns([0,0,0,1,1,0,0,0])
			psspy.fdns([0,0,0,1,1,0,0,0])
			psspy.fdns([0,0,0,1,1,0,0,0])
			psspy.fdns([0,0,0,1,1,0,0,0])
			psspy.fdns([0,0,0,1,1,0,0,0])
			k+=1
			step_name = 'Connect String' +str(String_num[m]) +'Segment'+str(WT_num[m])+'at Group'+str(i+1)
			get_results(step_name)

# Energization Over
columns = ['STEP','MAX_V_ONSSHV', 'MIN_V_ONSSHV', 'MAX_V_ONSSLV','MIN_V_ONSSLV', 'MAX_V_ONSRE_HDD_6','MIN_V_ONSRE_HDD_6','MAX_V_ONSRE_HDD_5','MIN_V_ONSRE_HDD_5','MAX_V_ONSRE_HDD_4','MIN_V_ONSRE_HDD_4','MAX_V_ONSRE_HDD_3','MIN_V_ONSRE_HDD_3','MAX_V_ONSRE_HDD_2','MIN_V_ONSRE_HDD_2','MAX_V_ONSRE_HDD_1','MIN_V_ONSRE_HDD_1','MAX_V_CNCRT_DB','MIN_V_CNCRT_DB','MAX_V_LNDFAL_HDD','MIN_V_LNDFAL_HDD','MAX_V_SEABED','MIN_V_SEABED', 'MAX_V_OFSSHV','MIN_V_OFSSHV','MAX_V_OFSSLV','MIN_V_OFSSLV']
# columns1 = [str(b) for b in ONSS_LV]
# columns1 = [str(b) for b in ONSS_LV]
columns += ['MAX_V_66kV','MIN_V_66kV','MAX_V_720V','MIN_V_720V','MAX_Q_WT','MIN_Q_WT','Qstat','Convergence Status']
print columns
# Store Data
df = pd.DataFrame(result_dict).T
df.columns = columns
df.to_excel(Results_file+r""".xlsx""")
