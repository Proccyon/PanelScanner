from PIL import Image
import numpy as np
import matplotlib.pyplot as plt
import xlwt
import sys,os



#-----GlobalFunctions-----#
def MakePath(FileName,FileType):
    return os.path.dirname(sys.argv[0])+r'/ '[0] +str(FileName)+"."+FileType

def MakeFolderPath(FolderName):
    return os.path.dirname(sys.argv[0])+r'\ '[0] +FolderName+r'\ '[0]
  
def FindImagePaths(FolderName="Panel_Input"):
    FolderPath = MakeFolderPath(FolderName)
    UrlList = os.listdir(FolderPath)
    
    NewUrlList = []
    NameList = []
    for Url in UrlList:
        if(".jpg" in Url or ".png" in Url):
            NewUrlList.append(FolderPath + r'\ '[0] +Url)
            NameList.append(Url)
    return NewUrlList,NameList
   
def ReadCounter():
    File = open(MakePath("Counter","txt"),"r")
    FileCounter = float(File.read())
    File.close()
    
    File = open(MakePath("Counter","txt"),"w")
    File.write(str(round(FileCounter+0.01,2)))
    File.close()
    return FileCounter     
              
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

#-----PanelSorter-----#

class PanelSorter:
    def __init__(self,Xmin,Xmax,PanelWidth,PanelPixels):
        self.Xmin = Xmin
        self.Xmax = Xmax
        self.PanelWidth = PanelWidth
        self.PanelPixels = PanelPixels
        
        self.CreatePanelList()
        self.WriteToExcel()
        
    def CreatePanelList(self):
        UrlList,self.NameList = FindImagePaths()
        
        self.PanelList = []
        for Url in UrlList:
            self.PanelList.append(Panel(Url,self.Xmin,self.Xmax,self.PanelWidth,self.PanelPixels))
            
    def WriteToExcel(self):
        
        Path = MakePath("PanelResults"+str(ReadCounter()),"xls")
    
        Book = xlwt.Workbook()
        Sheet1 = Book.add_sheet("MainResults")
        
        BoldStyle = xlwt.XFStyle()
        BoldStyle.bold = True
        
        Row0 = Sheet1.row(0)
        Row0.write(0,"Panel Name",style=BoldStyle)
        Row0.write(1,"Corrosion Distance",style=BoldStyle)
        
        Sheet1.col(0).width = 3500
        Sheet1.col(1).width = 3500
        
        for i in range(len(self.PanelList)):
            Sheet1.write(i+1,0,self.NameList[i])
            Sheet1.write(i+1,1,str(self.PanelList[i].CorrosionCM)+"cm")
            
        Book.save(Path)
        


#-----PanelSorter-----#




#-----Panel-----#
class Panel:
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
        plt.title("Blue filter of panel image")
        plt.imshow(self.ColorArray,cmap="Blues")
        plt.colorbar()
              
    def PrintResults(self):
        print("CorrosionDistance: "+str(self.CorrosionCM)+"cm")

    #Calculates corrosion from pixels to cm
    def ConvertCorrosion(self):
        return round(self.CorrosionPixels * self.LengthFactor,2)
    
#-----Panel-----#      
        
#-----Main-----#
Path = "C:/Users/Eigenaar/Desktop/PannelScanner/"
Name = "Pannel1.jpg"
URL = Path+Name

PanelWidth = 10.1#cm
PanelPixels = 2875
Xmin = 1500
Xmax = 2500

MainPanelSorter = PanelSorter(Xmin,Xmax,PanelWidth,PanelPixels)
plt.show()
#-----Main-----#
