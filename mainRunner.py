#all imports 
from os import getpid
from keras.models import load_model
from time import sleep
from keras.preprocessing.image import img_to_array
from keras.preprocessing import image
import cv2
import numpy as np
from GazeTracking.gaze_tracking import GazeTracking
import xlsxwriter     
import time      

#function to return the mapping for label to the values of the emotions
def get_number(label):
    if(label=="Angry"):
        return 0.86
    elif(label=="Disgust"):
        return 0.93
    elif(label=="Fear"):
        return 0.86
    elif(label=="Happy"):
        return 0.79
    elif(label == "Sad"):
        return -0.90
    elif(label=="Surprise"):
        return 0.85
    elif(label =="Neutral"):
        return  -0.82

webcam = cv2.VideoCapture(0) #open the camera

gaze = GazeTracking() #gaze tracking object
cascade_classifier=cv2.CascadeClassifier(r'C:\Users\Saptarushi\OneDrive\Desktop\MajorProject\EmotionRecognition\frontalface.xml') #model to search for face in an image
model=load_model(r'C:\Users\Saptarushi\OneDrive\Desktop\MajorProject\EmotionRecognition\face.hdf5') #emotion detection model
class_labels=['Angry','Disgust','Fear','Happy','Neutral','Sad','Surprise'] #classes of the emotions

#create an excel sheet to write write data
book = xlsxwriter.Workbook(r'C:\Users\Saptarushi\OneDrive\Desktop\MajorProject\Data\sheets\data.xlsx')     
sheet = book.add_worksheet('data')
row = 0    
column = 0   

#Time information variables
TT = 12 #runs for 1200 frames
sleep_time = 0.499 #sleeps for 0.499 ~ 0.5 secs, so runs at 2fps
t=0

i=0
while (TT > 0):
    _, frame = webcam.read()
    labels = [] #labels initialized as an empty list 

    gray = cv2.cvtColor(frame,cv2.COLOR_BGR2GRAY) # gray scale the image
    faces = cascade_classifier.detectMultiScale(gray,1.3,5) #detect the faces

    #for each face detected in the image
    for (x,y,w,h) in faces:

        cv2.rectangle(frame,(x,y),(x+w,y+h),(0,0,0),2) #put a rectagle on the face
        roi_gray = gray[y:y+h,x:x+w] #cut out the face in gray
        roi_gray = cv2.resize(roi_gray,(48,48),fx=100,fy=100,interpolation=cv2.INTER_CUBIC) #resize the face
        crop = frame[y:y+h,x:x+w] #save the face in color formate

        if np.sum([roi_gray])!=0: #is face is present, and the image is not all dark
            
            #EMotion Detection part
            roi = roi_gray.astype('float')/255.0
            roi = img_to_array(roi)
            roi = np.expand_dims(roi,axis=0)

            preds = model.predict(roi)[0] #do the prediction
            label=class_labels[preds.argmax()] #get the label
            label_position = (x,y)  #make a tuple of the coordinates where the face begins

            #Gaze Tracking part
            gaze.refresh(frame) #give the source image to the gaze tracking object
            frame = gaze.annotated_frame() #set the markings on the pupils
            hr=gaze.horizontal_ratio()
            vr=gaze.vertical_ratio()
            #label_position = (x-2,y-50)
            label = get_number(label)
            l=[]  
            l.append(hr)
            l.append(vr)
            l.append(get_number(class_labels[preds.argmax()]))
            l.append(preds.argmax())
            l.append(class_labels[preds.argmax()] )
            column = 0
            for x in l:
                print()
                sheet.write(row, column, x)
                column+=1
            
            cv2.imshow("Demo", frame)
            cv2.imwrite(r'C:\Users\Saptarushi\OneDrive\Desktop\MajorProject\Data\Images\kang'+str(i)+'.jpg',crop) #print the image into images folder
            sheet.insert_image(row,column,r'C:\Users\Saptarushi\OneDrive\Desktop\MajorProject\Data\Images\kang'+str(i)+'.jpg')
            
            print("image " + str(i) + "cap")
            
            row += 1          
            i+=1
    time.sleep(sleep_time) #sleeps till the amount of time specified in 'sleep_time' variable
    TT = TT -1 #reduce the number of frames captured
book.close()