from PIL import Image
import numpy as np
import matplotlib.pyplot as plt
import xlwt
import sys,os



#-----GlobalFunctions-----#
def MakePath(FileName,FileType):
    return os.path.dirname(sys.argv[0])+r'/ '[0] +FileName+"."+"FileType"

#-----GlobalFunctions-----#


#-----Pannel-----#
class Pannel:
    def __init__(self,URL,Xmin,Xmax,PanelWidth,PanelPixels,Axis=0,ColorName="Blue",FitDegree=13):
        self.URL = URL
        self.Image, self.R,self.G,self.B = self.GetImage()
        self.Xmin,self.Xmax = Xmin,Xmax
        self.Axis = Axis
        self.ColorName = ColorName
        self.FitDegree = FitDegree
        self.LengthFactor = PanelWidth / PanelPixels #Pixels / cm
        
    #Gets an image from the given path
    #Returns an array of pixels and R,G,B separately
    def GetImage(self):
        RawImage = Image.open(self.URL)
        Array = np.asarray(RawImage)
        return Array,Array[:,:,0],Array[:,:,1],Array[:,:,2]

    #Normalizes and crops the image
    def TransformArray(self,ColorArray):
        return (1- ColorArray / 255)[:,self.Xmin:self.Xmax]

    #Fits the ColourProfile to a polynomial to the corrosion distance
    #Returns Distance to the Left,Right and Total distance
    #Also plots the fit and minima
    def ColorFit(self,x,y,PlotFit=False):
        FitParameters = np.polyfit(x,y,self.FitDegree)
        ColorFit = np.poly1d(FitParameters)(x)
        
        LeftMin = self.FindLocalMin(ColorFit,int((self.Xmin+self.Xmax)/2-self.Xmin),-1)
        RightMin = self.FindLocalMin(ColorFit,int((self.Xmin+self.Xmax)/2-self.Xmin),+1)
        Corrosion = RightMin - LeftMin
        
        if(PlotFit):
            plt.plot(x,ColorFit,linestyle="--",label="Polynomial fit",color="black",alpha=1)
        plt.axvline(x=LeftMin,linestyle="--",color="red",label="Corrosion distance",alpha=0.7)
        plt.axvline(x=RightMin,linestyle="--",color="red")
        
        return LeftMin,RightMin,Corrosion
        
    #Titel, etc for Colour profile plot
    def SetupPlot1(self):
        
        plt.figure(figsize=(7,7))
        
        plt.subplot(1,2,1)
        
        plt.title("Average amount of "+self.ColorName+" over position")
        plt.xlabel("x(pixels)")
        plt.ylabel("Amount of "+self.ColorName)
        
        plt.xlim(xmin=0,xmax=self.Xmax-self.Xmin)
        
    #Titel, etc for Image plot
    def SetupPlot2(self):
        plt.subplot(1,2,2)
        plt.title("Blue filter of pannel image")

    def IsMinimum(self,y0,yLeft,yRight):
        return yLeft > y0 and yRight>y0

    #Finds the first local minimum going from Xstart to right/left
    def FindLocalMin(self,Array,Xstart,Direction=1):
        if(Direction==1):
            Xend = len(Array)
        else:
            Xend = 0
            Direction=-1
        
        for i in np.arange(Xstart,Xend-Direction,step=Direction):
            y0 = Array[i]
            yLeft = Array[i-1]
            yRight = Array[i+1]
            if(self.IsMinimum(y0,yLeft,yRight)):
                return i
        
        return None
        
    def PrintResults(self):
        print("CorrosionDistance: "+str(self.CorrosionCM)+"cm")

    #Calculates corrosion from pixels to cm
    def CalcCorrosion(self):
        return round(self.CorrosionPixels * self.LengthFactor,2)
        
    #Finds Corrosion distance and plots the colour profile and Image
    def MakeColorProfile(self):
        ColorArray = self.TransformArray(self.B)
        
        self.ColorProfile = np.average(ColorArray,axis=self.Axis)
        x = np.arange(len(self.ColorProfile))        
        
        self.SetupPlot1()
        plt.plot(x,self.ColorProfile,label="Color Average")
        self.LeftMin,self.RightMin,self.CorrosionPixels = self.ColorFit(x,self.ColorProfile)
        self.CorrosionCM = self.CalcCorrosion()
        plt.legend()
        
        self.SetupPlot2()
        plt.imshow(ColorArray,cmap="Blues")
        plt.colorbar()

        self.PrintResults()
        
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
Pannel1.MakeColorProfile()

plt.show()
#-----Main-----#