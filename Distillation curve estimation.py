'''
Nate Dee
MATH 300 Project 2
WSU-TC Fall 2026


AI Usage disclaimer:
    
    No AI tools were used in any way to compile the code, report, or poster for this project.

System requirements: 
    
    * This program was built and tested using Spyder 6 on a Windows 10 machine with
    Microsoft Excel installed. It has not been tested on any other combination of hardware. 
    
    * In order for this program to function, the appropriate excel file must be located in the
    same folder as this program. 
      

'''

#===============<<< PACKAGE IMPORT >>>===============
import numpy as np
import scipy.interpolate as si
import matplotlib.pyplot as plt
import xlwings as xw
#===============<<< MAIN CONTROLL >>>===============

def main():

    s1_Data, s2_Data, s3_Data, s4_Data = import_Workbook()   #Data pulled from workbook. 
    
    s1x, s1y = pull_Data(s1_Data)
    s2x, s2y = pull_Data(s2_Data)
    s3x, s3y = pull_Data(s3_Data)
    s4x, s4y = pull_Data(s4_Data)
    
    data_index = [(s1x, s1y), (s2x, s2y), (s3x, s3y), (s4x, s4y)]
     
    trg = [10, 50, 90]      # If a shape error occurs, increase the range for the target
    x_range = [10, 10, 10]
    
    print()
    print('================<<< FULL CURVE ANALYSIS >>>==================')
    print()
    
    for i in range(len(data_index)): # each sample
    
        print('Sample', (i+1), 'target analysis with full datasets:')
        print()
        print('        Linear    PCHIP    Akima   Makima')   
    
        for j in range(len(trg)): #each target

            make_table(data_index[i][0], data_index[i][1], trg[j])
            
        #plot_Data_Each(data_index[i][0], data_index[i][1], (i+1), trg, 1001)    #Produces 4 charts. Only activate if needed. Last input = interp resolution
        dipcheck(data_index[i][0], data_index[i][1], 1001) #Last input = interp resolution
        print()
         
    print()
    print('================<<< ZOOMED ANALYSIS >>>==================')
    print()
    
    for i in range(len(data_index)): # each sample
    
        print('Sample', (i+1), 'target analysis with zoomed datasets:')
        print()
        print('        Linear    PCHIP    Akima   Makima')  
    
        for j in range(len(trg)): #each target

            zx, zy = zoom(data_index[i][0], data_index[i][1], trg[j], x_range[j])
            make_table(zx, zy, trg[j])
            dipcheck(zx, zy, 101)
            
            #plot_Zoom_Each(zx, zy, (i+1), trg[j], x_range[j], 101)      #Produces 12 Charts. Only activate if needed. Last input = interp resolution
        
        print()
        
    print()
         
    return


#===============<<< DATA EXTRACTION FUNCTIONS >>>===============

def import_Workbook():  # Uses xlwings to open the excel file and copies cell values to a list.
                        # Extraction can't be done as a loop due to the datasets between samples not being the same length.
    
    xw.App(visible=False)
    
    filename = '_Draft3_ME436_R1_GCxGC_Boiling_Points.xlsx'
    sheet1 = 'S01_BP_Sorted'
    sheet2 = 'S02_BP_Sorted'
    sheet3 = 'S03_BP_Sorted'
    sheet4 = 'S04_BP_Sorted'
    
    wb = xw.Book(filename)
    
    s1_Imp = wb.sheets[sheet1]
    s2_Imp = wb.sheets[sheet2]
    s3_Imp = wb.sheets[sheet3]
    s4_Imp = wb.sheets[sheet4]
      
    s1_Data = s1_Imp.range('F8:H45').value # Cell A8:F45 for full range on s1   
    s2_Data = s2_Imp.range('F8:H43').value # Cell A8:F43 for full range on s2 
    s3_Data = s3_Imp.range('F8:H102').value # Cell A8:F102 for full range on s3
    s4_Data = s4_Imp.range('F8:H44').value # Cell A8:F44 for full range on s4
    
    wb.app.quit() # Closes background excel program that was opened to get data. 
        
    return(s1_Data, s2_Data, s3_Data, s4_Data)

#==========================    
def pull_Data(target):   # Extracts needed data and rounds mass percent to 3 decimals and temp to 2.

    x = [0] # Adding the point 0,0 to every dataset for interpolation methods
    y = [0] # because one of the targets is below the first data point. 

    for i in range(len(target)):
        x.append(target[i][2])
        y.append(target[i][0])
        
    return(x, y)

#==========================
def zoom(x, y, t, r): 
   
   zoomX = []
   zoomY = []
    
   for i in range((len(x))):
       if x[i] >= (t-r):
           if x[i] <= (t+r):
               zoomX.append(x[i])
               zoomY.append(y[i])         
   
   return(zoomX, zoomY)

#===============<<< DATA ANALYSIS FUNCTIONS >>>===============
def make_table(x, y, t):
        
    methods = [lambda z: np.interp(z, x, y),
           si.PchipInterpolator(x, y),
           si.Akima1DInterpolator(x, y, extrapolate=True),
           si.Akima1DInterpolator(x, y, method='makima', extrapolate=True)]
     
    #======>>> Sub function to make values align better in terminal <<<======
    def adjust(n):
        
        if n <= -100:
            n = '{:.1f}'.format(round(n, 2))
        else: 
            if -100 < n < 0:
                n = '{:.2f}'.format(round(n, 2))
            else: 
                if 0 <= n < 10:
                    n = '{:.4f}'.format(round(n, 2))
                else: 
                    if 10 <= n < 100 :
                        n = '{:.3f}'.format(round(n, 2))
                    else: 
                        if n >= 100:
                            n = '{:.2f}'.format(round(n, 2))
                            
        return(n)
    #===============================>>> <<<==================================
        
    lint   =  adjust((methods[0](t)))
    pchip  =  adjust(np.float64((methods[1](t))))
    akima  =  adjust(np.float64((methods[2](t))))
    makima =  adjust(np.float64((methods[3](t))))
     
    print('T ',t, ' ', lint, ' ', pchip, ' ', akima,' ', makima)
    
    pass

#==========================
def plot_Data_Each(x, y, n, t, res):     
       
    xnew = np.linspace(min(x),max(x),res)
    names = ['linear', 'PCHIP', 'Akima', 'Makima']
    methods = [lambda z: np.interp(z, x, y),
           si.PchipInterpolator(x, y),
           si.Akima1DInterpolator(x, y, extrapolate=True),
           si.Akima1DInterpolator(x, y, method='makima', extrapolate=True)]
    colors = ['red', 'blue', 'green', 'magenta']
     
    fig, axs = plt.subplots(figsize=(15, 7))
    axs.plot(x, y, label='Test Data', marker='o', color='tab:orange')
    for name, p, color in zip(names, methods, colors): 
        axs.plot(xnew, p(xnew), label=name, color=color)
        
    axs.set_xlabel('Percent Distilled')
    axs.set_ylabel('Temperature C') 
    for i in range(len(t)):
        axs.axvline(t[i], label= (f'T{t[i]}'), ls= '--', color= 'red', alpha= .5)
    axs.grid(True, alpha=0.3)
    axs.legend()     
    fig.suptitle(f'Distilation curve for sample {n}', fontsize=14)
    plt.tight_layout()
    
    pass

#==========================
def plot_Zoom_Each(x, y, n, t, r, res): 
    
    rl = t - r
    ru = t + r
    xnew = np.linspace(rl,ru,res)
    ylim_l = y[0]
    ylim_u = y[(len(y)-1)]
     
    names = ['linear', 'PCHIP', 'Akima', 'Makima']
    methods = [lambda z: np.interp(z, x, y),
           si.PchipInterpolator(x, y),
           si.Akima1DInterpolator(x, y, extrapolate=True),
           si.Akima1DInterpolator(x, y, method='makima', extrapolate=True)]
    colors = ['red', 'blue', 'green', 'magenta']

    fig, axs = plt.subplots(figsize=(15, 7))
    axs.set_xlim(rl, ru)
    axs.set_ylim(ylim_l - 5, ylim_u + 5)
    axs.axvline(t, label= f'T{t}', ls= '--', color= 'red', alpha= .5)
 
    axs.plot(x, y, label='Test Data', marker='o', color='tab:orange')
    
    for name, p, color in zip(names, methods, colors):
        axs.plot(xnew, p(xnew), label=name, color=color)

    axs.set_xlabel('Percent Distilled')
    axs.set_ylabel('Temperature C')
    axs.grid(True, alpha=0.3)
    axs.legend()
    
    fig.suptitle(f'Zoom section, sample {n}, T{t}', fontsize=14)
    plt.tight_layout()
   
    pass

#==========================
def dipcheck(x, y, res):
    
    xnew = np.linspace(min(x),max(x),res)
    methods = [lambda z: np.interp(z, x, y),
           si.PchipInterpolator(x, y),
           si.Akima1DInterpolator(x, y, extrapolate=True),
           si.Akima1DInterpolator(x, y, method='makima', extrapolate=True)] 

    lin = methods[0](xnew)
    pchip = methods[1](xnew)
    akima = methods[2](xnew)
    makima = methods[3](xnew)      
    
    def value(m): 
       
       v = 'null'
       
       for i in range(1, len(m)):
       
           if m[i] >= m[i-1]:
               v = 'Pass'
           else: 
               v = 'Fail'
               break
       
       return(v)
    
    ltest = value(lin)
    ptest = value(pchip)
    atest = value(akima)
    mtest = value(makima)         
    
    print('Dip     ', ltest, '    ', ptest, '   ', atest, '   ', mtest)

    pass

#===============================================
#===============<<< GO BUTTON >>>===============

main()
