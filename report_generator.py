import pandas as pd
import numpy as np
import xlsxwriter as xls
import math
import matplotlib.pyplot as plt
def get_number(label):
    if(label==0):
        return 0.788
    elif(label==1):
        return 0.511
    elif(label==2):
        return 0.788
    elif(label==3):
        return 0.847
    elif(label == 4):
        return 0.7083
    elif(label==5):
        return 0.725
    elif(label ==6):
        return  1
    else:
        return 0
def _process_sheet():
    standard_frames = 300
    #column numbers
    c0 = 0
    c1 = 1
    c2 = 2
    c3 = 3
    c4 = 4
    row_counter = 0

    #sheet variables
    sheet = None
    hr = []
    vr = []
    y_axis = [0,0,0,0,0,0,0]
    x_axis = [0,1,2,3,4,5,6]
    emo = []
    #read the excel sheet as a data frame
    with pd.ExcelFile(r'C:\Users\majaa\OneDrive\Documents\MajorProject-main\Data\sheets\data.xlsx') as reader:
        sheet = pd.read_excel(reader, sheet_name = 'data')
    data = sheet.to_numpy() #convert to np array
    #strip the values to horizontal and vertical to individual lists
    np.array(data[2], dtype=int)
    data[2]=np.array(data[2], dtype=np.int)
    new_array = data[:,2].tolist()
    np.ascontiguousarray(new_array)
    #[int(new_array) for new_array in new_array]
    #print(type(new_array))
   #print(new_array)
    for i in data:
        hr.append(i[0])
        vr.append(i[1])
        emo.append(get_number(i[2]))
    for i in range(len(new_array)):
        y_axis[int(new_array[i])]+=1

    #lists are ready
    #print(hr)
   # print(vr)
    print(y_axis)
    wb = xls.Workbook(r'hrvr_table_processed.xls',{'nan_inf_to_errors': True})
    sheet = wb.add_worksheet('hrvr_table')
    frames = 1
    average_hr = 0
    average_vr = 0
    average_emo = 0
    nan_emo = 0
    sum_hr = 0
    sum_vr = 0
    sum_emo = 0
    nan_s = 0
    for i in range(len(hr)):
        if(frames == standard_frames ):
            if(nan_s > 100): # more than 1/3 of time was spent undetected, so deem the frame unecessary
                sheet.write(row_counter,c3,-1)
                sheet.write(row_counter,c4,-1)
                sheet.write(row_counter,6,"Not Attentive")
                sheet.write(row_counter,5,"Not Attentive")
            else:
            #calculate averages
            #log the averages in to the excel sheet
                average_hr = sum_hr / (frames - (nan_s) )#ignore the number of time the model did not detect of person not found
                average_vr = sum_vr / (frames - (nan_s) ) # fix this bug later
                average_emo = sum_emo / (frames - (nan_s))
                if(average_hr>0.5 and average_hr<0.75):
                    sheet.write(row_counter,5,"Attentive")
                    print(average_hr)
                elif(average_hr==-1):
                    sheet.write(row_counter,5,"Not Attentive")
                else:
                    sheet.write(row_counter,5,"Not Attentive")
                    print(average_hr) 
                if(average_vr>0.6 and average_vr<0.85):
                    sheet.write(row_counter,6,"Attentive")
                    print(average_vr)
                elif(average_vr==-1):
                    sheet.write(row_counter,6,"Not Attentive")
                else:
                    sheet.write(row_counter,6,"Not Attentive")
                    print(average_vr)
                sheet.write(row_counter,c3,average_hr)
                sheet.write(row_counter,c4,average_vr)
                sheet.write(row_counter,7,average_emo)
                #print(sum_hr)
                #print(sum_vr)
            row_counter = row_counter + 1 #standard_frames
            nan_s = 0
            sum_hr = 0
            sum_vr = 0
            sum_emo = 0
            nan_emo = 0
            frames = 1
        if(math.isnan(hr[i])):
            nan_s= nan_s + 1
        else:
            print(emo[i],frames,nan_emo)
            sum_hr = sum_hr + hr[i]
            sum_vr = sum_vr + vr[i]
            sum_emo = sum_emo + emo[i]
            #print("sum_hr=",sum_hr)
            #print("sum_vr=",sum_vr)
        if(math.isnan(emo[i])):
            nan_emo = nan_emo + 1
        frames = frames + 1
    #log the rest of the data after the loop ends, might have less than 120 frames captured
    if(frames > 40):
        average_hr = sum_hr / (frames - (nan_s)) 
        average_vr = sum_hr / (frames - (nan_s)) 
        print(average_hr)
        sheet.write(row_counter,c3,average_hr)
        sheet.write(row_counter,c4,average_vr)
    plt.plot(x_axis,y_axis)
    plt.show()
    wb.close()
_process_sheet()