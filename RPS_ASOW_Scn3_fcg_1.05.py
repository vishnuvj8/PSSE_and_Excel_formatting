# Based on Caroline Janelle's code for the 66kV collector study
# Which is based on Asim's code OP1_Cable_Ampacity_Full_Network_01.py for the original Mayflower Reactive Power Study (105 WTs, 1 nm in between)
# Asim's code modified for new model QP829_Full_STATCOM_size106WT.sav
# Careful when modifying, code not clean

import os,sys
import math
import xlrd
from xlrd import open_workbook
import xlwt
import itertools
import pandas as pd
from openpyxl import load_workbook
import numpy

# sys.path.insert(0,'C:\\Program Files (x86)\\PTI\\PSSE33\\PSSBIN')
# os.environ['PATH'] = 'C:\\Program Files (x86)\\PTI\\PSSE33\\PSSBIN'+';'+os.environ['PATH']
# import redirect
# redirect.psse2py()
import psspy
import excelpy
# psspy.psseinit(0)
# _i=psspy.getdefaultint()
# _f=psspy.getdefaultreal()
# _s=psspy.getdefaultchar()


# Functions added to get max of voltage values in Full Collector System
# Functions 
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
    # if all "None", return "None"
    if all(isinstance(v, str) for v in _voltages):
        max_voltages = "None"
    # if some "None" and some values, return max of values
    else:
        max_voltages = max(v for v in _voltages if isinstance(v,float))
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
        max_q = max(q for q in _q if isinstance(q,float))
    return max_q    

def get_Q_statcom():
    ierr, Q_s1 = psspy.macdat(statcom1_bus ,'1','Q')
    ierr, Q_s2 = psspy.macdat(statcom2_bus ,'1','Q')
    if Q_s1 < 0 and Q_s2 < 0:
        Q_s = min(Q_s1,Q_s2)
    elif Q_s1 > 0 and Q_s2 > 0:        
        Q_s = max(Q_s1,Q_s2)
    return Q_s, Q_s1, Q_s2


#################################################INPUTS#####################################################
# Model of collector added to get max of voltage values in Full Collector System
# Swing buses
swing_bus = [206294]
POI_bus =swing_bus[0]
V_POI = 1.05
record_output = 1
psspy.case(r"""Larrabee_MW_MVOW_0.95.sav""")
ACF = 0.5
myxls = excelpy.workbook("RPS_ASOW_Scn3_13p6MW_"+str(V_POI)+".xlsx", sheet="RPS", overwritesheet=True)
num_shunt = 8 
oltc_ratio_ons = 0.96
oltc_ratio_ofs = 0.97
STATCOM_size = 50  
# statcom start value and iteration vset step
#Vset_STAT = .9497#0.9996#1.0087   ##############INPUT
#Vset_step = 0.000001        ##############INPUT
#Vset_STAT = 0.9996#1.0087   ##############INPUT
#Vset_step = 0.000001        ##############INPUT  
Vset_STAT = 1.0497#0.9996#1.0087   ##############INPUT
Vset_step = 0.000001        ##############INPUT
# Vset_STAT = 0.949947000000007 #special
loop_num = 650    
##########################################################################################################
# BUSES NEED TO BE ADAPTED FOR EACH SCENARIO
# 66 kV buses
offshore_txf_buses = [10401,10402]
wtstr_1 = [101, 102, 103, 104, 105, 106]
wtstr_2 = [201, 202, 203, 204, 205, 206]
wtstr_3 = [301, 302, 303, 304, 305, 306]
wtstr_4 = [401, 402, 403, 404, 405, 406]
wtstr_5 = [501, 502, 503, 504, 505, 506]
wtstr_6 = [601, 602, 603, 604, 605, 606]
wtstr_7 = [701, 702, 703, 704, 705, 706]
wtstr_8 = [801, 802, 803, 804, 805, 806]
wtstr_9 = [901, 902, 903, 904, 905, 906]
wtstr_10 = [1001, 1002, 1003, 1004, 1005, 1006]

#statcom buses
statcom1_bus = 10101
statcom2_bus = 10102
ons_reac1_bus = 10201
ons_reac2_bus = 10202


ofs_cable_1 = [10501,10601,10301,90101,10701,10801,10901,101001,101101,101201,10201]
ofs_cable_2 = [10502,10602,10302,90102,10702,10802,10902,101002,101102,101202,10202]


# WT strings to each offshore txf
wtstr_to_txf_1 = [wtstr_1,wtstr_2,wtstr_3,wtstr_4,wtstr_5]
wtstr_to_txf_2 = [wtstr_6,wtstr_7,wtstr_8,wtstr_9,wtstr_10]


# All 0.72 kV buses (66 kV bus numbers + 200)
buses_p72kV_temp_1 = wtstr_to_txf_1 + wtstr_to_txf_2
buses_p72kV_temp_2 = list(itertools.chain.from_iterable(buses_p72kV_temp_1))
buses_p72kV = [x*10 for x in buses_p72kV_temp_2]

# All 66 kV buses
buses_66kV_temp = [offshore_txf_buses] + wtstr_to_txf_1 + wtstr_to_txf_2
buses_66kV = list(itertools.chain.from_iterable(buses_66kV_temp))

# All 230 kV offshore cable buses
buses_ofs_230kV_Cable_temp = [ofs_cable_1,ofs_cable_2]
buses_ofs_230kV_Cable = list(itertools.chain.from_iterable(buses_ofs_230kV_Cable_temp))

psspy.solution_parameters_4([_i,_i,_i,_i,10],[_f,_f,_f,_f, ACF,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])

Vset_STAT_Arr = []
V_OnHV = []
V_OnLV = []


V_ofs_cab_230kV_min = []
V_ofs_cab_230kV_max = []

V_66kV_min = []
V_66kV_max = []
V_p72kV_min = []
V_p72kV_max = []
I_land = []
I_cncr_duct_1 = []
I_cncr_duct_2 = []
I_cncr_duct_3 = []
I_cncr_duct_4 = []
I_onsh_1 = []
I_onsh_2 = []
I_onsh_3 = []
I_lnd_HDD = []
I_Sea = []
I_Jtube = []
P_POI = []
Q_POI = []
PF = []
Q_stat1 = []
Q_stat2 = []
#Fxd_shnt = []
Tr_ONS_1 = []
Tr_ONS_2 = []
Tr_OFS_1 = []
Tr_OFS_2 = []
Qs_WTs_min = [] # UPDATE_Q_WT
Qs_WTs_max = [] # UPDATE_Q_WT

Num_Reactors = []
Convergence_test = []

ONS_TF = [] #UPDATE_XFO_LOADING
OFS_TF = [] #UPDATE_XFO_LOADING


# switch off shunts initially
psspy.plant_data(POI_bus,0,[V_POI, 100.0])
for i_sh in range (1,num_shunt+1):
    psspy.shunt_chng(ons_reac1_bus,str(i_sh),0,[_f,_f]) 
    psspy.shunt_chng(ons_reac2_bus,str(i_sh),0,[_f,_f]) 
    psspy.fdns([1,0,0,1,1,0,0,0])
    psspy.fdns([1,0,0,1,1,0,0,0])


psspy.plant_data(statcom1_bus,_i,[Vset_STAT,_f])
psspy.plant_data(statcom2_bus,_i,[Vset_STAT,_f])
# psspy.plant_data(128,_i,[Vset_STAT,_f])

psspy.fdns([1,0,0,1,1,0,0,0])
psspy.fdns([1,0,0,1,1,0,0,0])


#########################################################################################
Converged = 1
count = 1
while (Converged == 1):
    psspy.fdns([1,0,0,1,1,0,0,0])
    psspy.fdns([1,0,0,1,1,0,0,0])
    psspy.fnsl([1,0,0,1,1,0,0,0])
    psspy.fnsl([1,0,0,1,1,0,0,0])
    psspy.fnsl([1,0,0,1,1,0,0,0])
    psspy.fnsl([1,0,0,1,1,0,0,0])
    psspy.fnsl([1,0,0,1,1,0,0,0])
    Converged = psspy.solved()
    count = count + 1
    if count > 5:
        print ( "Case is non-convergent.")
        time.sleep(10)
        sys.exit()



Column_Lables = ['Vset_STAT (pu)','V_OnHV (pu)','V_OnLV (pu)','Tr_ONS_1','Tr_ONS_2','Tr_OFS_1','Tr_OFS_2','I_land (%)','I_cncr_duct_1 (%)','I_onsh_1 (%)','I_cncr_duct_2 (%)','I_onsh_2 (%)','I_cncr_duct_3 (%)','I_onsh_3 (%)','I_cncr_duct_4 (%)','I_lnd_HDD (%)','I_Sea (%)','I_Jtube','ONS_XFO (%)','OFS_XFO (%)','P_POI (MW)','Q_POI (MVar)','PF','Cnvg.','Reactors (nos.)','Q_stat1_final (MVar)','Q_stat2_final (MVar)','V_ofs_cab_min (PU)','V_ofs_cab_max (PU)','V_66kV_min (PU)','V_66kV_max (PU)','V_0.72_kV_min (PU)','V_0.72_kV_max (PU)','Qs_WTs_min (MVAr)','Qs_WTs_max (MVAr)'] # UPDATE_Q_WT # UPDATE_XFO_LOADING


myxls.set_range(1,1,[Column_Lables],
                transpose=False,fontStyle='bold', fontName=None, fontSize=None, fontColor=None, wrapText=False,
                numberFormat=None, sheet="RPS")

if record_output == 1:
    psspy.lines_per_page_one_device(1,1000000)
    psspy.progress_output(2,r"""output_record.txt""",[0,0])

count = 0   
for it in range (0,loop_num): 
    print "abcdefg"
    print it
    psspy.plant_data(statcom1_bus,_i,[Vset_STAT,_f])
    psspy.plant_data(statcom2_bus,_i,[Vset_STAT,_f])
    # psspy.plant_data(128,_i,[Vset_STAT,_f])
    psspy.fdns([1,0,0,1,1,0,0,0])
    psspy.fnsl([1,0,0,1,1,0,0,0])
    psspy.fdns([1,0,0,1,1,0,0,0])
    psspy.fnsl([1,0,0,1,1,0,0,0])
    psspy.fdns([1,0,0,1,1,0,0,0])
    psspy.fnsl([1,0,0,1,1,0,0,0])
    psspy.fdns([1,0,0,1,1,0,0,0])
    Converged = psspy.solved()
    Convergence_test.append(Converged)
    if Converged == 1:
        count = count + 1

    Q_s = get_Q_statcom()

    #Q_stat_0.append(Q_s)
    i_sh = 1

    for k in range (1,num_shunt+1):   
        if Q_s[0] <= -STATCOM_size:    
            psspy.shunt_chng(ons_reac1_bus,str(i_sh),1,[_f,_f]) 
            psspy.shunt_chng(ons_reac2_bus,str(i_sh),1,[_f,_f]) 
            # psspy.shunt_chng(128,str(i_sh),1,[_f,_f])
            ##onshore transformers
            psspy.two_winding_chng_6(10101,10201,r"""1""",[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f, oltc_ratio_ons,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)
            psspy.two_winding_chng_6(10102,10202,r"""1""",[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f, oltc_ratio_ons,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)

            ##offshore transformers
            psspy.two_winding_chng_6(10501,10401,r"""1""",[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f, oltc_ratio_ofs,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)
            psspy.two_winding_chng_6(10502,10402,r"""1""",[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f, oltc_ratio_ofs,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s,_s)

            psspy.fdns([1,0,0,1,0,0,0,0])
            psspy.fdns([1,0,0,1,0,0,0,0])
            psspy.fdns([1,0,0,1,0,0,0,0])
            Q_s = get_Q_statcom() 
            i_sh = i_sh + 1       
    
        else:
            print ('breaking')
            break
    
    psspy.fdns([1,0,0,1,0,0,0,0])
    psspy.fdns([1,0,0,1,1,0,0,0])
    psspy.fdns([1,0,0,1,1,0,0,0])
    psspy.fdns([1,0,0,1,1,0,0,0])

    # ierr, fs = psspy.busdt1(108,'YS','ACT')   
    Num_Reactors.append(i_sh-1)        
    Q_stat1.append(Q_s[1])
    Q_stat2.append(Q_s[2])

    
    ierr, P = psspy.macdat(POI_bus ,'1','P')
    P_POI.append(abs(P))

    ierr, Q = psspy.macdat(POI_bus ,'1','Q')
    Q_POI.append(Q)

    ierr, S = psspy.macdat(POI_bus ,'1','MVA')

    # ierr, P = psspy.brnmsc(110,109,'1','P')
    # P_POI.append(abs(P))
    # ierr, Q = psspy.brnmsc(110,109,'1','Q')
    # Q_POI.append(Q)
    # ierr, S = psspy.brnmsc(110,109,'1','MVA')
    PF.append(abs((P/S)))


    fbus = [POI_bus, 10202, 101202, 101102, 101002, 10902, 10802, 10702, 90102, 10302, 10602]
    tbus = [10102,  101202, 101102, 101002, 10902, 10802, 10702, 90102, 10302, 10602, 10502]

    br = 0
    # Used to check 'AMPS' instead of '%'
    ierr, I1 = psspy.brnmsc(fbus[br],tbus[br],'1','PCTRTA')
    ierr, I2 = psspy.brnmsc(tbus[br],fbus[br],'1','PCTRTA')
    I_land.append(max(I1,I2))

    br = 1
    ierr, I1 = psspy.brnmsc(fbus[br],tbus[br],'1','PCTRTA')
    ierr, I2 = psspy.brnmsc(tbus[br],fbus[br],'1','PCTRTA')
    I_cncr_duct_1.append(max(I1,I2))

    br = 2
    ierr, I1 = psspy.brnmsc(fbus[br],tbus[br],'1','PCTRTA')
    ierr, I2 = psspy.brnmsc(tbus[br],fbus[br],'1','PCTRTA')
    I_onsh_1.append(max(I1,I2))
        
    br = 3
    ierr, I1 = psspy.brnmsc(fbus[br],tbus[br],'1','PCTRTA')
    ierr, I2 = psspy.brnmsc(tbus[br],fbus[br],'1','PCTRTA')
    I_cncr_duct_2.append(max(I1,I2))

    br = 4
    ierr, I1 = psspy.brnmsc(fbus[br],tbus[br],'1','PCTRTA')
    ierr, I2 = psspy.brnmsc(tbus[br],fbus[br],'1','PCTRTA')
    I_onsh_2.append(max(I1,I2))

    br = 5
    ierr, I1 = psspy.brnmsc(fbus[br],tbus[br],'1','PCTRTA')
    ierr, I2 = psspy.brnmsc(tbus[br],fbus[br],'1','PCTRTA')
    I_cncr_duct_3.append(max(I1,I2))

    br = 6
    # Used to check 'AMPS' instead of '%'
    ierr, I1 = psspy.brnmsc(fbus[br],tbus[br],'1','PCTRTA')
    ierr, I2 = psspy.brnmsc(tbus[br],fbus[br],'1','PCTRTA')
    I_onsh_3.append(max(I1,I2))

    br = 7
    ierr, I1 = psspy.brnmsc(fbus[br],tbus[br],'1','PCTRTA')
    ierr, I2 = psspy.brnmsc(tbus[br],fbus[br],'1','PCTRTA')
    I_cncr_duct_4.append(max(I1,I2))

    br = 8
    ierr, I1 = psspy.brnmsc(fbus[br],tbus[br],'1','PCTRTA')
    ierr, I2 = psspy.brnmsc(tbus[br],fbus[br],'1','PCTRTA')
    I_lnd_HDD.append(max(I1,I2))
        
    br = 9
    ierr, I1 = psspy.brnmsc(fbus[br],tbus[br],'1','PCTRTA')
    ierr, I2 = psspy.brnmsc(tbus[br],fbus[br],'1','PCTRTA')
    I_Sea.append(max(I1,I2))

    br = 10
    ierr, I1 = psspy.brnmsc(fbus[br],tbus[br],'1','PCTRTA')
    ierr, I2 = psspy.brnmsc(tbus[br],fbus[br],'1','PCTRTA')
    I_Jtube.append(max(I1,I2))


###################################### - COPY FROM HERE - ##### #UPDATE_XFO_LOADING ###########################
    # ONS_xfo_loading_update
    xfo_fbus = [10101, 10102] 
    xfo_tbus = [10201, 10202]
    br = 0
    ierr, L1 = psspy.brnmsc(xfo_fbus[br],xfo_tbus[br],'1','PCTRTA')
    ierr, L2 = psspy.brnmsc(xfo_fbus[br],xfo_tbus[br],'1','PCTRTA')
    br = 1
    ierr, L3 = psspy.brnmsc(xfo_fbus[br],xfo_tbus[br],'1','PCTRTA')
    ierr, L4 = psspy.brnmsc(xfo_fbus[br],xfo_tbus[br],'1','PCTRTA')
    ONS_TF.append(max(L1,L2,L3,L4))


    # OFS_xfo_loading_update
    xfo_fbus = [10501, 10502]
    xfo_tbus = [10401, 10402]
    br = 0
    ierr, L1 = psspy.brnmsc(xfo_fbus[br],xfo_tbus[br],'1','PCTRTA')
    ierr, L2 = psspy.brnmsc(xfo_fbus[br],xfo_tbus[br],'1','PCTRTA')
    br = 1
    ierr, L3 = psspy.brnmsc(xfo_fbus[br],xfo_tbus[br],'1','PCTRTA')
    ierr, L4 = psspy.brnmsc(xfo_fbus[br],xfo_tbus[br],'1','PCTRTA')
    OFS_TF.append(max(L1,L2,L3,L4))
    
###################################### - TO HERE - ##### #UPDATE_XFO_LOADING ###########################


    ierr, V_bus_OnHV = psspy.busdat(10102 ,'PU')
    ierr, V_bus_OnLV = psspy.busdat(10202 ,'PU')
    # ierr, V_bus110 = psspy.busdat(110 ,'PU')
    V_OnHV.append(V_bus_OnHV)
    V_OnLV.append(V_bus_OnLV)
    # V_110.append(V_bus110)

    # Added Max and min voltages of collector system, and added all ofshore cable buses
    res_voltages_ofs_cab_230kV = read_voltages(buses_ofs_230kV_Cable)
    res_voltages_66kV = read_voltages(buses_66kV)
    res_voltages_p72kV = read_voltages(buses_p72kV)
    res_Qs_WTs = read_mach_Q(buses_p72kV) # UPDATE_Q_WT

    V_ofs_cab_230kV_min.append(min(res_voltages_ofs_cab_230kV))
    V_ofs_cab_230kV_max.append(max(res_voltages_ofs_cab_230kV))
    V_66kV_min.append(min(res_voltages_66kV))
    V_66kV_max.append(max_voltages(res_voltages_66kV))
    V_p72kV_min.append(min(res_voltages_p72kV))
    V_p72kV_max.append(max_voltages(res_voltages_p72kV))
    Qs_WTs_min.append(min(res_Qs_WTs)) # UPDATE_Q_WT
    Qs_WTs_max.append(max_q(res_Qs_WTs)) # UPDATE_Q_WT

    # Removed, aggregated collector not in model anymore
    # ierr, V_bus = psspy.busdat(301 ,'PU')
    # V_301.append(V_bus)    

    ierr, tr = psspy.xfrdat(10201, 10101, '1', 'RATIO')
    Tr_ONS_1.append(tr)
    ierr, tr = psspy.xfrdat(10202, 10102, '1', 'RATIO')
    Tr_ONS_2.append(tr)

    ierr, tr = psspy.xfrdat(10401, 10501, '1', 'RATIO')
    Tr_OFS_1.append(tr)
    ierr, tr = psspy.xfrdat(10402, 10502, '1', 'RATIO')
    Tr_OFS_2.append(tr)

    Vset_STAT_Arr.append(Vset_STAT)
    Vset_STAT = Vset_STAT + Vset_step

    

    # switching off shunts at the end
    for i_sh in range (1,num_shunt+1):
        psspy.shunt_chng(ons_reac1_bus,str(i_sh),0,[_f,_f]) 
        psspy.shunt_chng(ons_reac2_bus,str(i_sh),0,[_f,_f]) 
        psspy.fdns([1,0,0,1,1,0,0,0])
        psspy.fdns([1,0,0,1,1,0,0,0])    

psspy.progress_output(1,r"""output_record.txt""",[0,0])
myxls.set_range(2,1,[Vset_STAT_Arr,V_OnHV,V_OnLV,Tr_ONS_1,Tr_ONS_2,Tr_OFS_1,Tr_OFS_2,I_land,I_cncr_duct_1,I_onsh_1,I_cncr_duct_2,I_onsh_2,I_cncr_duct_3,I_onsh_3,I_cncr_duct_4,I_lnd_HDD,I_Sea,I_Jtube,ONS_TF,OFS_TF,P_POI,Q_POI,PF,Convergence_test,Num_Reactors,
                Q_stat1,Q_stat2,V_ofs_cab_230kV_min,V_ofs_cab_230kV_max,V_66kV_min,V_66kV_max,V_p72kV_min,V_p72kV_max,Qs_WTs_min,Qs_WTs_max],
                transpose=True,fontStyle='regular', fontName=None, fontSize=None, fontColor=None, wrapText=False,
                numberFormat=None, sheet="RPS") # UPDATE_Q_WT #UPDATE_XFO_LOADING
print('saving excel')
myxls.save()

myxls.close()

print ("End of code. "+str(count)+" cases did not converge.")

