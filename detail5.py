from comtypes.client import CreateObject
from comtypes.client import GetActiveObject
import array

#Create Autocad file
try:
    Aapp= GetActiveObject("AutoCad.Application")
    Aapp.Visible= True
    Adrawing=Aapp.ActiveDocument
except:
    Aapp = CreateObject("AutoCad.Application")
    Aapp.Visible = True
    Adrawing = Aapp.ActiveDocument
ms= Adrawing.Modelspace

from comtypes.gen.AutoCAD import *

from pyautocad import Autocad, APoint


acad = Autocad()
acad.prompt("Hello, Autocad from Python\n")
    












#Takes a list of points and creates a list of distances between subsequent
#points
def dist(ent):
    a=[]
    from math import sqrt
    for i in range(1,len(ent)):
        #starting from the second element, it creates distnce between it and
        #its previous points
        firs=ent[i]
        sec=ent[i-1]
        #pythagorean theorem
        ans=sqrt((firs[0]-sec[0])**2+(firs[1]-sec[1])**2)
        a.append(ans)
    return(a)


#Takes a list of vertices and creates a list of midpoints between each location
#and the one before it
def mid_points(ent):
    a=[]
    from math import sqrt
    for i in range(1,len(ent)):
        firs=ent[i]
        sec=ent[i-1]
        ans=[(firs[0]+sec[0])/2,(firs[1]+sec[1])/2]
        a.append(ans)
    return(a)

#This is linear function. It takes overall length and uses modulus operator.
#a stock length and lap length are provided. It takes the total length and returns
#the quantity of stock lengths needed and the remaining rebar at the end
def linear(leng,barl,lapl,clearance=.25):
    leng=leng-2*clearance
    
    quant=int(leng/(barl-lapl))
    rem=leng%(barl-lapl)
    if leng<=barl+2:
        return([leng])

        
    if rem<=lapl+2:
        quant=quant-1
        rem=barl-(lapl-rem)
    
    
    
    if quant==0:
        return([rem])
    if leng<lapl:
        return([leng])
    
    
    return([quant,rem])

#Takes a length in feet and returns architectural units (feet'-inches")
def imperial(number):
    feet=int(number)
    inches=int(round((number-int(number))*12))
    if inches==12:
        feet=feet+1
        inches=0
    return(str(feet)+'-'+str(inches))


#Takes rebar, a list containing reinforcement sections. Uses length of the edge
#and reinforcement of that edge. Int the case of continuous rebars, uses the linear
#function and returns a string containing quantity of stock rebar. For spaced
#rebar, it also returns quantity of that rebar with its label
def estimate(rebar,length,barl,lap=3):
    from math import ceil
    text=''
    for i in rebar:
        if isinstance(i[0], int):
            a=linear(length,barl,lap)
            quant=i[0]
            if len(a)==2:
                quant=(i[0])*a[0]
            size=i[1]
            notes=i[2]
            if len(a)==2:
                text=text+(str(quant)+""+size+" "+imperial(barl)+ " "+notes+" "+str(i[0])+" Runs")+'\n'
                text=text+(str(i[0])+""+size+" "+imperial(a[1])+" Lap "+str(int(lap*12))+'"')+'\n'
            else:
                text=text+(str(quant)+""+size+" "+imperial(a[0])+ " "+notes+" "+str(i[0])+" Runs")+'\n' 
        else:
            spacing=i[2]/12
            quant=str(ceil(length/spacing)+1)
            size=i[0]
            length2=str(i[1])
            notes=i[3]
            text=text+(quant+""+size+" "+length2+ " "+notes+"@"+str(i[2]))+'\n'
    return(text)
      
    
#This one takes in a string and returns a list. It takes string with standard
#detailing notation for reinforcement. It returns a list for rebar. For continous
#rebar it takes something similar to "4#5 CONT Footing". It returns a list of
#length 3 ([4 ,"#5", CONT Footing]) For spaced rebar it will be similar to
# "#5 MK501 DWLS@48" it will return of length 4 [#5, "MK501", 48, DWLS]
def to_list(x):
    if x[0]=='#':
        x1=x
        x1=x1.replace(' ',',')
        x1=x1.replace('@',',')
        x1=x1.replace('"',"")
        x1=x1.split(',')
        x2=[]
        x2.append(x1[0])
        x2.append(x1[1])
        x2.append(int(x1[-1]))
        x3=x1[:-1]
        x2.append(1)
        x2[3]=' '.join(x3[2:])
        return(x2)
    else:
        x1=x
        x1=x1.replace(' ',',')
        x1=x1.replace('#',',#')
        x1=x1.split(',')
        x1[0]=int(x1[0])
        x1[2]=' '.join(x1[2:])
        x1=x1[0:3]
        return(x1)
    







#This function takes two points in 2-D Space and returns the distance between
def normal(a,b):
    distance=((a[0]-b[0])**2+(a[1]-b[1])**2)**.5
    return distance




def aselection():
    x=acad.get_selection(text='Select objects')
    y=[]
    for i in x:
        y.append(list(i.Coordinates))
    z=[]
    for i in range(len(y)):
        q=[]
        for j in range(len(y[i])):
            if j%2==0:
                q.append([y[i][j]/12,y[i][j+1]/12])
        z.append(q)
    y2=[]
    for i in x:
        y2.append(str(i.Color))
    z2=[z,y2]
    return(z2)














def print_detail(vertices,width,rebar,height):
    global count
    try: count
    except NameError: count = None

    if count==None:
        count=1
    rebar2=[]
    for i in range(len(rebar)):
        rebar2.append(to_list(rebar[i]))
    rebar=rebar2
    
    
    
    
    
    
    print(acad.doc.Name)
 


    from matplotlib import pyplot as plt
    distances=dist(vertices)
    midpoints=mid_points(vertices)
    for j,i in enumerate(distances):
        print(count)

        
        
        s=estimate(rebar,i+width,30)
        if normal(vertices[0],vertices[-1])>.05:
            r=(j==0)
            t=(j==(len(distances)-1))
            sums=r+t
            s=estimate(rebar,i+width-sums*width/2,30)
            k=i+width-sums*width/2







        
        [x1,y1]=midpoints[j]
        p1=APoint(12*x1,12*y1)
        s2=s.split('\n')
        colors = ['r', 'g', 'b', 'm','c','y']
        if i>5:
            for kn in s2:
                text = acad.model.AddText(kn, p1, height)
                text.color = count%6+1
                p1.y -= 1.5*height
                
            
        else:
            plt.text(x1,y1,str(count),size=0.2,verticalalignment="center",horizontalalignment='center',color='k')
            s=s.replace('\n','\nxyz')
            s='xyz'+s
        print(s)
        print(' ')
        print(' ')
        print(' ')
        count=count+1
    
    def counter():
        return count








def make_detail(filename,footing_width,rebar,color,height):
    from matplotlib import pyplot as plt
    count=1
    [vertices,colors]=aselection()
    
    colorss={'red':'1','yellow':'2','green':'3','cyan':'4','blue':'5','magenta':'6'}
    for j,i in enumerate(vertices):
        color_num=colors[j]
        
        if colorss[color]==color_num:
            print_detail(i,footing_width,rebar,height)
        x,y= map(list, zip(*i))
        plt.plot(x,y)
        plt.axis('equal') 
    fname='C:/Users/mosman/Documents/detail2.pdf'
    plt.savefig(fname)
    
