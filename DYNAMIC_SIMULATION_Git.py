import os
import sys
sys_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE33\PSSBIN'
sys.path.append(sys_path_PSSE)
env_path_PSSE=r' C:\Program Files (x86)\PTI\PSSE33\PSSBIN'
os.environ['PATH'] += ';' + env_path_PSSE

import psspy
import redirect
import dyntools

pssbindir  = r"C:\Users\Directory..."
savfile    = "C:\Users\Directory...\CASE.sav"
snpfile    = r"C:\Users\Directory...\python_test.snp"
outfile    = r"C:\Users\Directory...\python_test.out"

# Redirect output from PSSE to python
redirect.psse2py()

# Open the Saved Case (.../.sav Data)
CASE = r"C:\Users\Directory...\CASE.sav"
psspy.psseinit(12000)
psspy.case(CASE)

# Convert Loads and Generators ( Which is itself a 3 step process):
psspy.conl(-1,1,1)
psspy.conl(-1,1,1,[0,0],[100,0,0,100])
psspy.conl(-1,1,3)

# Convert Generators
psspy.cong()

# Solve for dynamics 
psspy.ordr()
psspy.fact()
psspy.tysl()

# Save converted case ( The new case is saved where there is no need to make the convert again)
case_root = os.path.splitext(CASE)[0]
psspy.save(case_root + "Converted_Case.sav")

# Add Dynamics data file and load it 
psspy.dyre_new(dyrefile= "C:\Users\Directory...\dynamic.dyr")

#psspy.change_plmod_con(1,r"""1""",r"""GENSAL""",4, 9.54)
#psspy.change_plmod_con(2,r"""1""",r"""GENROU""",5, 3.34)

# Add channels and parameters 
# Bus Voltage
# psspy.chsb(sid=0,all=1, status=[-1,-1,-1,1,13,0])
# Active and Reactive Power Flow
psspy.chsb(sid=0,all=1, status=[-1,-1,-1,1,7,0])
psspy.chsb(sid=0,all=1, status=[-1,-1,-1,1,12,0])
psspy.chsb(sid=0,all=1, status=[-1,-1,-1,1,14,0])
psspy.chsb(sid=0,all=1, status=[-1,-1,-1,1,16,0])

# Set simulation parameter step size
psspy.dynamics_solution_params(realar3=0.001)

# Save snapshot
psspy.snap(sfile="C:\Users\Directory...\python_test.snp")

# Read data #############################################
import excelpy
xlsfile = r"data.xlsx"
print "Reading Data"
print "Data read from Excel %s file.\n"% xlsfile
xlsdata = excelpy.workbook(xlsfile,"test",mode='r')
row=1
col=1
print xlsdata.get_cell((row,col))
rowe = row+10
print xlsdata.get_cell((row,col))
datalst = xlsdata.get_range((row,col,rowe,col+2))
for x in datalst: print x

# Initialize the case 
psspy.strt(outfile="C:\Users\Directory...\python_test.out")

for n in range(row,rowe):
    #print xlsdata.get_cell((1+n,2))
    #print xlsdata.get_cell((2+n,3))
    time = xlsdata.get_cell((2+n-1,2))
    power = xlsdata.get_cell((2+n-1,3))
    time = float(time)
    power = float(power)
    if power > 1:
        power = 1
    #power = xlsdata.get_cell((1+n,3))
    #psspy.change_wnmod_var(1,r"""1""",r"""WT4E1""",4, 0.8)
    psspy.run(0, time, 1, 1, 0) # Time = second variable
    psspy.change_wnmod_var(1,r"""1""",r"""WT4E1""",4, power)
    
def test_data_extraction(chnfobj):   
    print '\n Testing call to get_data'
    sh_ttl, ch_id, ch_data = chnfobj.get_data()
    print sh_ttl
    print ch_id

    print '\n Testing call to get_id'
    sh_ttl, ch_id = chnfobj.get_id()
    print sh_ttl
    print ch_id
    
    print '\n Testing call to get_range'
    ch_range = chnfobj.get_range()
    print ch_range

    print '\n Testing call to get_scale'
    ch_scale = chnfobj.get_scale()
    print ch_scale

    print '\n Testing call to print_scale'
    chnfobj.print_scale()

    print '\n Testing call to txtout'
    chnfobj.txtout(channels=[])

    print '\n Testing call to xlsout'
    chnfobj.xlsout(channels=[1,2,3,4,5,7,9,10])

def test1():

    # create out files, run this only once and comment out once you have the .out files
    # create object
   
    outlst = [outfile]
    chnf = dyntools.CHNF(outlst)
    test_data_extraction(chnf)
    
if __name__ == '__main__':

    # Need to run test1 before running test2, test3 or test4
    # After running test1, you need to run test2, test3 and test4 one at a time.
    test1()