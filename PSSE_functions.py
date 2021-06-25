
from __future__ import division
import pandas as pd
import os,sys
import numpy
import math
import xlwt
import itertools
import csv

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
#new
#import psse33
#import psspyomp as psspy
psspy.number_threads(8)

########################################################################################################################

def run_power_flow():
    #ierr = psspy.fdns([1,0,0,1,1,0,0,0])
    #ierr = psspy.fdns([1,0,0,1,1,0,0,0])
    #ierr = psspy.fdns([1,0,1,1,1,0,0,0])
    
    ierr = psspy.fdns([0,0,0,1,1,0,0,0])
    ierr = psspy.fdns([1,0,0,1,1,0,0,0])
    ierr = psspy.fdns([1,0,1,1,1,0,0,0])
       
    
    return ierr

#returns an array: [P, Pmin, Pmax, Q, Qmin, Qmax] (in this order)
def gen_data(bus_nbr, gen_id):
    temp = ['P','PMIN', 'PMAX', 'QMIN', 'QMAX']
    array_results = []
    
    for i in range(len(temp)):
        ierr, temp_value = psspy.macdat(bus_nbr, str(gen_id), temp[i])
        if ierr == 0:
            array_results.append(temp_value)
    
    return array_results

def scale_gen(MW_change, list_redispatched_gen):
    
    # Define a new subsystem (call it nbr 3) --all gen that will be redispatched
    psspy.bsys(1,0,[0.0,0.0],0,[],len(list_redispatched_gen),list_redispatched_gen,0,[],0,[])
    psspy.scal_2(1,0,1,[0,0,0,0,0],[0.0,0.0,0.0,0.0,0.0,0.0,0.0])
    psspy.scal_2(0,1,2,[_i,3,1,4,0],[0.0,MW_change,0.0,0.0,0.0,0.0, 0.95])


    return

#returns the MW margin we have on a given set of generators
def allowed_gen_scale_margin(df_redispatched_gen):
    delta_gen_change_down = 0
  
    for index, row in df_redispatched_gen.iterrows():
        
        gen_bus = row["Bus number"]
        gen_id = row["Id"]
        #read gen data
        temp = gen_data(gen_bus, gen_id)

        #actual loading of gen
        loading_gen_MW = temp[0]
        #min_gen
        min_gen_MW = temp[1]
        max_gen_MW = temp[2]
      
        delta_gen_change_down = delta_gen_change_down + (loading_gen_MW - min_gen_MW)

    return  delta_gen_change_down

#returns the MW margin we have on a given set of generators
def disconnect_gens(df_disconnect_gen):
    delta_gen_change_down = 0
  
    for index, row in df_disconnect_gen.iterrows():
        
        gen_bus = row["Bus number"]
        gen_id = row["Id"]
        #disconnect gen        
        psspy.machine_chng_2(gen_bus,str(gen_id),[0,_i,_i,_i,_i,0],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f, 1.0])

    return 


def change_gen_output(bus_nbr, gen_id,new_P_out):
    psspy.machine_chng_2(bus_nbr,str(gen_id),[_i,_i,_i,_i,_i,_i],[ new_P_out,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    return

def add_new_gen(gen_bus, P_inj, Q_max, Q_min):
    psspy.bus_data_3(gen_bus,[1,1,1,1],[0.0, 1.0,0.0, 1.1, 0.9, 1.1, 0.9],"")
    psspy.bus_chng_3(gen_bus,[2,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f],_s)
    psspy.plant_data(gen_bus,0,[ 1.0, 100.0])
    psspy.machine_data_2(gen_bus,r"""1""",[1,1,0,0,0,0],[0.0,0.0, 9999.0,-9999.0, 9999.0,-9999.0, 100.0,0.0, 1.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0])
    psspy.machine_chng_2(gen_bus,r"""1""",[_i,_i,_i,_i,_i,1],[ P_inj,_f, Q_max,Q_min, P_inj,0.0,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    
    return


def add_2wind_txf(fr_bus, to_bus, br_id,br_owner,R,X,B,Rate_A, Rate_B,Rate_C):
    psspy.two_winding_data_4(fr_bus,to_bus,str(br_id),[1,to_bus,br_owner,0,0,0,33,0,fr_bus,0,1,0,1,1,1],[ R, X,100.0, 1.0,0.0,0.0, 1.0,0.0, Rate_A, Rate_B, Rate_C,1.0, 1.0, 1.0, 1.0,0.0,0.0, 1.1, 0.9, 1.1, 0.9,0.0,0.0,0.0],["",""])
    
    return 

def add_branch(fr_bus, to_bus, br_id,br_owner,R,X,B,Rate_A, Rate_B,Rate_C):

    psspy.branch_data(fr_bus, to_bus,str(br_id),[1,fr_bus,br_owner,0,0,0],[ R, X, B, Rate_A, Rate_B, Rate_C,0.0,0.0,0.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0])
    
    return
     
    
def add_dummy_branch(fr_bus, to_bus, br_id):
    
    psspy.branch_data(int(fr_bus),int(to_bus),str(br_id),[1,int(fr_bus),1,0,0,0],[0.0, 0.0001,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0])
    
    return

def save_case(file_name, folder_name):
    psspy.save(folder_name + '/' + file_name + ".sav")
    
    return


def load_case(path_to_case, case_name):
    #load the appropriate base case
    Case = path_to_case + case_name
    ##Load the SAV file
    psspy.case(Case)
    #increase iteration limit to 200
    psspy.solution_parameters_4([_i,200,_i,_i,10],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    #solve load flow
    #Check that the load flow converged (if it converged: ierr = 0)
    ierr = run_power_flow()

    return

def change_PAR_MW_limits(fr_bus,to_bus,br_id,v_max,v_min):
    psspy.two_winding_chng_4(fr_bus,to_bus,str(br_id),[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f, _f,_f, v_max,v_min,_f,_f,_f],["",""])    
    
    return


#performs an ACCC analysis and generate overload and non-convergence reports in two separate text files
def run_accc(path_to_case, case_name, sub_file,mon_file,con_file,subsystem_name,accc_outfile, overload_summary_file, nonconvergence_summary_file,accc_folder_results):
    
    #os.chdir(accc_folder_results)

    load_case(path_to_case, case_name)
    
    psspy.dfax_2([1,1,0],sub_file,mon_file,con_file,accc_outfile)

    #psspy.accc_with_dsp_3( 0.5,[0,0,0,1,1,2,0,0,0,0,0],subsystem_name,accc_outfile + ".dfx",accc_outfile,"","","")

    psspy.accc_parallel_2( 0.5,[0,0,0,1,1,2,0,0,0,0,0],subsystem_name,accc_outfile + ".dfx",accc_outfile,"","","")

    #generate overload report
    psspy.report_output(2,accc_folder_results + overload_summary_file + ".txt",0)
    psspy.accc_single_run_report_4([0,1,2,1,1,0,1,0,0,0,0,0],[0,0,0,0,6000],[ 0.5, 5.0, 100.0,0.0,0.0,0.0, 99999.],accc_outfile)
    
    #generate non-convergence report
    psspy.report_output(2,accc_folder_results + nonconvergence_summary_file + ".txt",0)
    psspy.accc_single_run_report_4([5,1,2,1,1,0,1,0,0,0,0,0],[0,0,0,0,6000],[ 0.5, 5.0, 100.0,0.0,0.0,0.0, 99999.],accc_outfile)

    return

#reads an overload ACCC output file and extract relevant info
def truncate_text_overloads(original_text_file, directory):
    os.chdir(directory)
    start = "<----------------- MONITORED BRANCH -----------------> <----- CONTINGENCY LABEL ------>   RATING     FLOW       %"
    end = " MONITORED VOLTAGE REPORT:"
    buffer = ""
    write_line = False
    for line in open(original_text_file):
      if line.strip() == start.strip():
        buffer = line
        write_line = True
      elif line.strip() == end.strip():
        write_line = False
      elif write_line:
        buffer += line
    new_file_name = original_text_file[:-4] + '_truncated.txt'
    open(new_file_name, 'w').write(buffer)
    if write_line == True:
      print("End string was not found -- overload file")
    return new_file_name

#reads a nonconvergent ACCC output file and extract relevant info
def truncate_text_nonconvergence(original_text_file, directory):
    os.chdir(directory)
    start = "X----- CONTINGENCY LABEL ------X X-- BUS ---X X- SYSTEM -X TERMINATION CONDITION"
    end = "CONTINGENCY LEGEND:"
    buffer = ""
    write_line = False
    for line in open(original_text_file):
      if line.strip() == start.strip():
        buffer = line
        write_line = True
      elif line.strip() == end.strip():
        write_line = False
      elif write_line:
        buffer += line
    new_file_name = original_text_file[:-4] + '_truncated.txt'
    open(new_file_name, 'w').write(buffer)
    if write_line == True:
      print("End string was not found -- nonconvergence file")
    return new_file_name

def read_and_extract_accc_overload(accc_results_directory):

    files = os.listdir(accc_results_directory)

    #for filename in accc_results_directory:
    for filename in files:
        
        if ".txt" in filename and "overload_report" in filename:
            new_file_name = truncate_text_overloads(filename, accc_results_directory)
            
            _widths = [6,13,7,6,13,7,3,36,9,9,5]
            data = pd.read_fwf(new_file_name, widths = _widths, skiprows = 1, skipfooter = 1,
                        names = ['from bus nbr', 'from bus name', 'from bus kV', 'to bus nbr', 'to bus name', 'to bus kV', 'ID', 'Contingency', 'Rating', 'Flow', 'Percent'])
            
            data.to_csv(filename[:-4] + '.csv') 


#data_base is the csv file containing the overloads pertaining to the reference base case
#data_new is the csv file containing the overloads pertaining to the case with the project
#folder is where all .csv file to be compared are located
def compare_overload(without_project_file, with_project_file, folder):

    results_files = []
    #read overload files
    #base case
    data_base = pd.read_csv(folder + without_project_file + ".csv")
    
    #modified (new) case
    data_new = pd.read_csv(folder + with_project_file + ".csv")

    overload_matrix_base = data_base.iloc[:,0:12].values

    overload_matrix_new = data_new.iloc[:,0:12].values

    list_of_changes = []

    #overload level beyond which we say there is an exacerbation of the situation 
    exacerbation_level = 3

    for i in range(len(overload_matrix_new)):
        found = 0

        for j in range(len(overload_matrix_base)):
             
            if (int(overload_matrix_new [i][1]) == int(overload_matrix_base [j][1])) and (int(overload_matrix_new [i][4]) == int(overload_matrix_base [j][4])) and (str(overload_matrix_new [i][7]) == str(overload_matrix_base [j][7])) and (str(overload_matrix_new [i][8]).strip() == str(overload_matrix_base [j][8]).strip()):
            
                found = 1

                #add to the list only if the overload has been exacerbated by more than exacerbation level
                
                if (overload_matrix_new [i][11] > overload_matrix_base [j][11] + exacerbation_level):
                    
                    temp = overload_matrix_new[i]
                    temp = map(str, temp) 
                    temp = temp + [str(overload_matrix_base [j][11])]
                    list_of_changes.append(temp)
                    

            #a new overload not present in the base case        
        if found == 0 and overload_matrix_new [i][11] >= 100 + exacerbation_level:
            #print(overload_matrix_new [i][10])
            
            temp = overload_matrix_new[i]
            temp = map(str, temp) 
            temp = temp + ["new overload"]
            list_of_changes.append(temp)
            

    df = pd.DataFrame(list_of_changes)

    if df.empty:
        df["Message"] = ["there are no new overload instances"]

    else:
        #delete first columns, which contains f. nothing
        try:
            df.drop(df.columns[[0]], axis=1, inplace=True)
        except:
            pass

        df.columns = ["from bus nbr","from bus name","from bus kV","to bus nbr","to bus name","to bus kV","ID","Contingency","Rating","Flow","Percent","Old percent"]
    
    # file_name_temp = "list_NEW_" + with_project_file + "_.xlsx"    
    # df.to_excel(file_name_temp,sheet_name = with_project_file)

    file_name_temp = "list_NEW_" + with_project_file + "_.csv"  
    df.to_csv(file_name_temp) 

    return 


def read_and_extract_accc_nonconvergence(accc_results_directory):

    files = os.listdir(accc_results_directory)

    #for filename in accc_results_directory:
    for filename in files:
        
        if ".txt" in filename and "nonconverge_report" in filename:
            #original_text_file = 'ACCC_results.txt'
            new_file_name = truncate_text_nonconvergence(filename, accc_results_directory)
            
            _widths = [32,13,13,25]
            data = pd.read_fwf(new_file_name, widths = _widths, skiprows = 1, skipfooter = 1,
                        names = ["contingency label","MW mismatch","MVAR mismatch","non convergence nature"])
            
            data.to_csv(filename[:-4] + '.csv') 

    return
    
#data_base is the csv file containing the list of nonconvergences pertaining to the reference base case
#data_new is the csv file containing the list of nonconvergences pertaining to the case with the project
#folder is where all .csv file to be compared are located
def compare_nonconvergence(without_project_file, with_project_file, folder):

    results_files = []
    #read overload files
    #base case
    data_base = pd.read_csv(folder + without_project_file + ".csv")
    
    #modified (new) case
    data_new = pd.read_csv(folder + with_project_file + ".csv")

    nonconvergence_matrix_base = data_base.iloc[:,0:5].values

    nonconvergence_matrix_new = data_new.iloc[:,0:5].values

    list_of_changes = []

    for i in range(len(nonconvergence_matrix_new)):
        found = 0

        for j in range(len(nonconvergence_matrix_base)):
             
            if ((nonconvergence_matrix_new [i][1]) == (nonconvergence_matrix_base [j][1])):
            
                found = 1

            #a new overload not present in the base case        
        if found == 0:            
            temp = nonconvergence_matrix_new[i]
            list_of_changes.append(temp)


    df = pd.DataFrame(list_of_changes)
    #delete first columns, which contains f. nothing

    if df.empty:
        df["Message"] = ["there are no new non-convergence instances"]

    else:
        try:
            df.drop(df.columns[[0]], axis=1, inplace=True)
        except:
            pass

        df.columns = ["contingency label","MW mismatch","MVAR mismatch","non convergence nature"]

    # file_name_temp = "list_NEW_" + with_project_file + "_.xlsx"            
    # df.to_excel(file_name_temp,sheet_name = with_project_file)

    file_name_temp = "list_NEW_" + with_project_file + "_.csv"  
    df.to_csv(file_name_temp) 

    return 