
from __future__ import division
import pandas as pd
import os,sys
import numpy as np
import math
import xlwt
import itertools
import PSSE_functions
import csv
from os import listdir
from openpyxl import load_workbook

########################################################################################################################

sys.path.insert(0,'C:\\Program Files (x86)\\PTI\\PSSE33\\PSSBIN')
os.environ['PATH'] = 'C:\\Program Files (x86)\\PTI\\PSSE33\\PSSBIN'+';'+os.environ['PATH']
import redirect
redirect.psse2py()
import psspy
psspy.psseinit(0)
_i=psspy.getdefaultint()
_f=psspy.getdefaultreal()
_s=psspy.getdefaultchar()


########################################################################################################################
########################################################################################################################
#INPUT
new_project_folder = "C:\Users\merit\Desktop\Projects\Equinor Assignments\Projects\Barrett injection study\\"

#path_to_case = 'C:\Users\merit\Desktop\Projects\EDF Assignment\Projects\SIS in Delaware (PJM)\PJM base cases/'
path_to_case = new_project_folder  + r"""Input\PSSE base case\\"""
#this should be the base case without the studied project but with queued projects and their corresponding generation redispatch

case_name_without_project = "Q958_Oceanside_Energy_2024SUM_OFF.sav"

sub_file = new_project_folder + r"""Input\My monsubcon\INJ_STD_NY.sub"""
mon_file = new_project_folder + r"""Input\My monsubcon\Combined NYISO and Mine.mon"""
con_file = new_project_folder + r"""Input\My monsubcon\INJ_STD_NY.con"""


accc_folder_results = accc_folder_results_compare = new_project_folder + r"""Working\Python\ACCC results\Compare max\\"""

#subsytem name, as defined in the .sub file
subsystem_name = "NY"

#list of dispatchable gens: you can specify two lists: one that is a priority list (usually composed of generators as close as possible to the point of renewable injection) and a more general one (with the rest of generators) 
list_redisp_gen_NYISO = pd.read_excel(new_project_folder + r"""Working\Python\List of scalable conventional gens in NYISO - Zone K.xlsx""", sheet_name="General")
list_redisp_gen_NYISO = list_redisp_gen_NYISO["Bus number"].values

priority_list_redisp_gen_NYISO = pd.read_excel(new_project_folder + r"""Working\Python\List of scalable conventional gens in NYISO - Zone K.xlsx""", sheet_name="Priority")
priority_list_redisp_gen_NYISO_array = priority_list_redisp_gen_NYISO["Bus number"].values

disconnect_list_redisp_gen_NYISO = pd.read_excel(new_project_folder + r"""Working\Python\List of scalable conventional gens in NYISO - Zone K.xlsx""", sheet_name="Disconnect")
disconnect_list_redisp_gen_NYISO_array = disconnect_list_redisp_gen_NYISO["Bus number"].values


########################################################################################################################
########################################################################################################################

#record psse output in a text file
record_output = 0

#do you want to generate the ACCC results for the base reference case (case without Project)?

#do you need to generate the PSSE files? --this is for the addition of the Project
#newly generated .sav files are saved in "path_to_case" 
#the case with Q'ed projects and without out project should be located in "path_to_case"

#do you want to run an ACCC analysis?
#results of accc analysis are saved in "accc_folder_results"
#the different .sav cases should all be located in "path_to_case"

#do you want to run the accc results comparison
#reads files located in folder "accc_folder_results_compare", which should contain the PSSE generated ACCC results (overload and nonconvergence); put both the nonconvergence and overloads of ref case (without project; this should be in csv format) and the other overload and nonconvergence files to be evaluated
#the results are saved in folder "accc_folder_results_compare"
#specify name of overload and nonconvergence of reference case
without_project_file_overloads = "ref_overload_report"
without_project_file_nonconvergence = "ref_nonconverge_report"

#specify the studied injection levels
injection_levels = [1200]
#injection_levels = [1000,1500]

#min and mac Q of gen -- in percent of P (injection_levels)
Qmax = 0.328
Qmin = 0.328

#POI info: (bus must exist in base case)
POI_bus = 129203

#in case you want to skip the simulation related to the generation of the ACCC for the ref case
skip_ref_accc = 0

#do you have a gen list for the generators that need to be disconnect? (1: Yes; 0: No)
disconnect_gen_list = 1

#do you have a priority list for the generators that need to be scaled? (1: Yes; 0: No)
priority_gen_list = 1

#specify if the used case is with upgrades?
case_with_upgrades = 1

#if set to 1, it will not generate a case for the with project scenario and will use the case already in the folder
skip_project_case_generation = 1
########################################################################################################################
#initial bus number of newly added gen bus
gen_bus = 300
gen_id = '1'
########################################################################################################################
#need to put extracted data in array format [a,b,c...] (currently extracted as [a b c]); although still correct, I cannot pass the extracted table (without commas) to a particular PSSE function (scale gens) as it requires the array element be separated by a comma
temp_1 = []
for i in range(len(list_redisp_gen_NYISO)):
    temp_1.append(list_redisp_gen_NYISO[i])
list_redisp_gen_NYISO = temp_1
temp_2 = []
if priority_gen_list:
    for i in range(len(priority_list_redisp_gen_NYISO_array)):
        temp_2.append(priority_list_redisp_gen_NYISO_array[i])
    priority_list_redisp_gen_NYISO_array = temp_2

temp_3 = []
if disconnect_gen_list:
    for i in range(len(disconnect_list_redisp_gen_NYISO_array)):
        temp_3.append(disconnect_list_redisp_gen_NYISO_array[i])

    disconnect_list_redisp_gen_NYISO_array = temp_3

########################################################################################################################
########################################################################################################################
execution_sequence = [0,0,0,0]

for i in range(len(execution_sequence)):
    
    if skip_ref_accc and i ==0:
        execution_sequence[i] = 0
    else:
        execution_sequence[i] = 1
                  
        
    run_accc_ref_case = execution_sequence[0]
    generate_sav_files = execution_sequence[1]
    run_accc = execution_sequence[2]
    results_comparison = execution_sequence[3]

    if record_output == 1:
        psspy.lines_per_page_one_device(1,1000000)
        psspy.progress_output(2,r"""output_record.txt""",[0,0])
        
    print(results_comparison + run_accc + generate_sav_files + run_accc_ref_case)

    if (results_comparison + run_accc + generate_sav_files + run_accc_ref_case) <= 1:
        if generate_sav_files and skip_project_case_generation == 0:
            for i in range(len(injection_levels)):

                PSSE_functions.load_case(path_to_case, case_name_without_project)
                PSSE_functions.run_power_flow()
                
                if disconnect_gen_list:
                    allowed_gen_scale_margin_disconnect = PSSE_functions.allowed_gen_scale_margin(disconnect_list_redisp_gen_NYISO)

                    if allowed_gen_scale_margin_disconnect >= injection_levels[i]:
                        #add gen
                        PSSE_functions.add_new_gen(gen_bus, injection_levels[i], injection_levels[i] * Qmax, - injection_levels[i] * Qmin) 
                        PSSE_functions.add_dummy_branch(gen_bus, POI_bus, 1)
                        #scale down gen
                        PSSE_functions.scale_gen(-(injection_levels[i]), disconnect_list_redisp_gen_NYISO_array)
                        PSSE_functions.run_power_flow()                        
                        
                    else:
                        #add gen
                        PSSE_functions.add_new_gen(gen_bus, allowed_gen_scale_margin_disconnect, allowed_gen_scale_margin_disconnect * Qmax, - allowed_gen_scale_margin_disconnect * Qmin) 
                        PSSE_functions.add_dummy_branch(gen_bus, POI_bus, 1)
                        #scale down gen
                        PSSE_functions.disconnect_gens(disconnect_list_redisp_gen_NYISO)
                        PSSE_functions.run_power_flow()
                        
                        if priority_gen_list:
                            #calculate the amount of gen MW change that can be accomplished using the priority list of gens; note: the dataframe should have two columns "Bus number" and "Id"
                            allowed_gen_scale_margin_priority = PSSE_functions.allowed_gen_scale_margin(priority_list_redisp_gen_NYISO)
                            
                            
                            if allowed_gen_scale_margin_priority >= (injection_levels[i] - allowed_gen_scale_margin_disconnect):
                                
                                PSSE_functions.change_gen_output(gen_bus, gen_id,injection_levels[i])
                                
                                PSSE_functions.scale_gen(-(injection_levels[i] - allowed_gen_scale_margin_disconnect), priority_list_redisp_gen_NYISO_array)
                                ierr = PSSE_functions.run_power_flow()
                                
                                Dummy_test = "I was there"
   
                            else:
                                PSSE_functions.change_gen_output(gen_bus, gen_id,allowed_gen_scale_margin_disconnect + allowed_gen_scale_margin_priority)
                                PSSE_functions.scale_gen(-allowed_gen_scale_margin_priority, priority_list_redisp_gen_NYISO_array)
                                ierr = PSSE_functions.run_power_flow()
                                
                                PSSE_functions.change_gen_output(gen_bus, gen_id,injection_levels[i])
                                PSSE_functions.scale_gen(-(injection_levels[i] - allowed_gen_scale_margin_priority - allowed_gen_scale_margin_disconnect), list_redisp_gen_NYISO)
                                ierr = PSSE_functions.run_power_flow()
                                
                        #no priority list        
                        else:                         
                            PSSE_functions.change_gen_output(gen_bus, gen_id,injection_levels[i])
                            PSSE_functions.scale_gen(-(injection_levels[i]-allowed_gen_scale_margin_disconnect), list_redisp_gen_NYISO)
                            ierr = PSSE_functions.run_power_flow()
                            
                        
                else:        
                
                    if priority_gen_list:
                        #calculate the amount of gen MW change that can be accomplished using the priority list of gens; note: the dataframe should have two columns "Bus number" and "Id"
                        allowed_gen_scale_margin_priority = PSSE_functions.allowed_gen_scale_margin(priority_list_redisp_gen_NYISO)                       
                        
                        
                        if allowed_gen_scale_margin_priority >= injection_levels[i]:
                            #add gen
                            PSSE_functions.add_new_gen(gen_bus, injection_levels[i], injection_levels[i] * Qmax, - injection_levels[i] * Qmin) 
                            PSSE_functions.add_dummy_branch(gen_bus, POI_bus, 1)
                            #scale down gen                            
                            PSSE_functions.scale_gen(-injection_levels[i], priority_list_redisp_gen_NYISO_array)
                            ierr = PSSE_functions.run_power_flow()
                        else:
                            #add gen
                            PSSE_functions.add_new_gen(gen_bus, allowed_gen_scale_margin_priority, allowed_gen_scale_margin_priority * Qmax, - allowed_gen_scale_margin_priority * Qmin) 
                            PSSE_functions.add_dummy_branch(gen_bus, POI_bus, 1)
                            
                            #scale down gen by allowed_gen_scale_margin_priority
                            PSSE_functions.scale_gen(-allowed_gen_scale_margin_priority, priority_list_redisp_gen_NYISO_array)
                            ierr = PSSE_functions.run_power_flow()
                            
                            # now use the list_redisp_gen_NYISO to scale down the remaining MW
                            #change gen MW
                            PSSE_functions.change_gen_output(gen_bus, gen_id,injection_levels[i])
                            #scale down gen
                            PSSE_functions.scale_gen(-(injection_levels[i] - allowed_gen_scale_margin_priority), list_redisp_gen_NYISO)
                            ierr = PSSE_functions.run_power_flow()                          
                            
                            

                    #no priority list        
                    else:                         
                        #add gen
                        PSSE_functions.add_new_gen(gen_bus, injection_levels[i], injection_levels[i] * Qmax, - injection_levels[i] * Qmin) 
                        PSSE_functions.add_dummy_branch(gen_bus, POI_bus, 1)
                        #scale down gen
                        PSSE_functions.scale_gen(-injection_levels[i], list_redisp_gen_NYISO)
                        ierr = PSSE_functions.run_power_flow()
                if case_with_upgrades:
                #here I connect the PAR (which was already in the base case -- it was just switched off)
                    psspy.branch_chng(401,129203,r"""1""",[0,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                    
                    #PAR limits of new PARs
                    v_max = 300
                    v_min = 290
                    #reconnect PARs
                    psspy.two_winding_chng_4(401,129203,r"""2""",[1,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],["",""])
                    psspy.two_winding_chng_4(401,129203,r"""3""",[1,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],["",""])
                    psspy.two_winding_chng_4(401,129203,r"""4""",[1,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],["",""])
                    
                    #change PAR MW limits
                    PSSE_functions.change_PAR_MW_limits(401,129203,r"""2""",v_max,v_min)
                    PSSE_functions.change_PAR_MW_limits(401,129203,r"""3""",v_max,v_min)
                    PSSE_functions.change_PAR_MW_limits(401,129203,r"""4""",v_max,v_min)
                    #ierr = PSSE_functions.run_power_flow()
                    
                    PSSE_functions.change_PAR_MW_limits(129203,129204,r"""1""",170,160)
                    ierr = PSSE_functions.run_power_flow()
        
                    #psspy.two_winding_chng_4(401,129203,r"""4""",[1,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],["",""])
                    ierr = PSSE_functions.run_power_flow()
                
                file_name = str(injection_levels[i]) + "MW"
                
                PSSE_functions.save_case(file_name, path_to_case)

                gen_bus = gen_bus + 1 

        #results of accc analysis are saved in "accc_folder_results"
        #the different .sav cases should all be located in "path_to_case"
        if run_accc or run_accc_ref_case:
            for i in range(len(injection_levels)):

                #run accc for ref case without Projetc
                if run_accc_ref_case:
                    overload_summary_file = "ref_overload_report"
                    nonconvergence_summary_file = "ref_nonconverge_report"
                    case_name = case_name_without_project
                    accc_outfile = "wihout_project_accc"

                else:
                    overload_summary_file = str(injection_levels[i]) + "_overload_report"
                    nonconvergence_summary_file = str(injection_levels[i]) + "_nonconverge_report"
                    case_name = str(injection_levels[i]) + "MW"
                    accc_outfile = str(injection_levels[i]) + "_accc"

                PSSE_functions.run_accc(path_to_case, case_name, sub_file,mon_file,con_file,subsystem_name,accc_outfile, overload_summary_file, nonconvergence_summary_file,accc_folder_results) 

        if results_comparison:
            #read overload file and extract relevant info --> save in .csv file
            PSSE_functions.read_and_extract_accc_overload(accc_folder_results_compare)
            #read nonconvergence file and extract relevant info --> save in .csv file
            PSSE_functions.read_and_extract_accc_nonconvergence(accc_folder_results_compare)

            #compare overloads
            files = os.listdir(accc_folder_results_compare)

            #for filename in accc_folder_results:
            for filename in files:
                
                if ".csv" in filename and "overload_report" in filename and "ref" not in filename:
                    PSSE_functions.compare_overload(without_project_file_overloads, filename[:-4], accc_folder_results_compare)
                
                if ".csv" in filename and "nonconverge_report" in filename and "ref" not in filename:
                    PSSE_functions.compare_nonconvergence(without_project_file_nonconvergence, filename[:-4], accc_folder_results_compare)


        #reset execution sequence
        execution_sequence = [0,0,0,0]
    else:
        print("#####################################################################")
        print("more than one simulation option selected. select only one at a time.")
        print("#####################################################################")
    
    files = os.listdir(accc_folder_results_compare)
    #deleted useless files
    for filename in files:
        if "truncated" in filename:
            print(filename)
            os.remove(filename)
        
