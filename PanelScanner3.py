from PIL import Image
import numpy as np
import matplotlib.pyplot as plt
import xlwt
import sys,os



#-----GlobalFunctions-----#
def MakePath(FileName,FileType):
    return os.path.dirname(sys.argv[0])+r'/ '[0] +FileName+"."+"FileType"

    #Gets an image from the given path
    #Returns an array of pixels and R,G,B separately
def GetImage(URL):
    RawImage = Image.open(URL)
    Array = np.asarray(RawImage)
    return Array,Array[:,:,0],Array[:,:,1],Array[:,:,2]
    
#Normalizes and crops an image
def TransformArray(Array,Xmin,Xmax):
    return (1- Array / 255)[:,Xmin:Xmax]
    
def IsMinimum(y0,yLeft,yRight):
    return yLeft > y0 and yRight>y0

#Finds the first local minimum going from Xstart to right/left
def FindLocalMin(Array,Xstart,Direction=1):
    if(Direction==1):
        Xend = len(Array)
    else:
        Xend = 0
        Direction=-1
        
    for i in np.arange(Xstart,Xend-Direction,step=Direction):
        y0 = Array[i]
        yLeft = Array[i-1]
        yRight = Array[i+1]
        if(IsMinimum(y0,yLeft,yRight)):
            return i
        
    return None
#-----GlobalFunctions-----#


#-----Pannel-----#
class Pannel:
    def __init__(self,URL,Xmin,Xmax,PanelWidth,PanelPixels,Axis=0,ColorName="Blue",FitDegree=13,DoPlot=False):
        self.URL = URL
        self.Image, self.R,self.G,self.B = GetImage(self.URL)
        self.Xmin,self.Xmax = Xmin,Xmax
        self.Axis = Axis
        self.ColorName = ColorName
        self.FitDegree = FitDegree
        self.LengthFactor = PanelWidth / PanelPixels #Pixels / cm
        
        self.DoColorProfileMethod()
        if(DoPlot):
            self.PlotResults()
      
    def DoColorProfileMethod(self):
        
        #Makes the color profile
        self.ColorArray = TransformArray(self.B,self.Xmin,self.Xmax)
        self.ColorProfile = np.average(self.ColorArray,axis=self.Axis)
        self.x = np.arange(len(self.ColorProfile))    
        
        #Fits a polynomial to the profile
        FitParameters = np.polyfit(self.x,self.ColorProfile,self.FitDegree)
        self.ColorFit = np.poly1d(FitParameters)(self.x)
        
        #Finds the local minima of the fit
        self.LeftMin = FindLocalMin(self.ColorFit,int((self.Xmin+self.Xmax)/2-self.Xmin),-1)
        self.RightMin = FindLocalMin(self.ColorFit,int((self.Xmin+self.Xmax)/2-self.Xmin),+1)
        
        #Calculates the corrosion
        self.CorrosionPixels = self.RightMin - self.LeftMin
        self.CorrosionCM = self.ConvertCorrosion()
    
    def PlotResults(self):
        
        #Color Profile Plot
        plt.figure(figsize=(7,7))
        
        plt.subplot(1,2,1)
        
        plt.title("Average amount of "+self.ColorName+" over position")
        plt.xlabel("x(pixels)")
        plt.ylabel("Amount of "+self.ColorName)
        
        plt.xlim(xmin=0,xmax=self.Xmax-self.Xmin)
        
        plt.plot(self.x,self.ColorProfile,label="Color Average")
        plt.plot(self.x,self.ColorFit,linestyle="--",label="Polynomial fit",color="black",alpha=1)
        plt.axvline(x=self.LeftMin,linestyle="--",color="red",label="Corrosion distance",alpha=0.7)
        plt.axvline(x=self.RightMin,linestyle="--",color="red")
        
        plt.legend()
        
        #Image plot
        plt.subplot(1,2,2)
        plt.title("Blue filter of pannel image")
        plt.imshow(self.ColorArray,cmap="Blues")
        plt.colorbar()
              
    def PrintResults(self):
        print("CorrosionDistance: "+str(self.CorrosionCM)+"cm")

    #Calculates corrosion from pixels to cm
    def ConvertCorrosion(self):
        return round(self.CorrosionPixels * self.LengthFactor,2)
    
    def ExportToExcel(self):
        Path = MakePath("Results","xls")
    
        Book = xlwt.Workbook()
        Sheet1 = Book.add_sheet("MainResults")
        
        BoldStyle = xlwt.XFStyle()
        BoldStyle.bold = True
        
        Row0 = Sheet1.row(0)
        Row0.write(0,"Pannel Name",style=BoldStyle)
        Row0.write(1,"Corrosion Distance",style=BoldStyle)
        
        Row1 = Sheet1.row(1)
        Row1.write(0,Name)
        Row1.write(1,self.CorrosionCM)
        
        
  
#-----Pannel-----#      
        
    

#-----Main-----#
Path = "C:/Users/Eigenaar/Desktop/PannelScanner/"
Name = "Pannel1.jpg"
URL = Path+Name

PannelWidth = 10.1#cm
PannelPixels = 2875
Xmin = 1500
Xmax = 2500

Pannel1 = Pannel(URL,Xmin,Xmax,PannelWidth,PannelPixels)

plt.show()
#-----Main-----#
