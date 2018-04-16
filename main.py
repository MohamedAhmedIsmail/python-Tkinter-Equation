import pandas as pd
import tkinter as tk
import math
class GUI:
    def __init__(self):
        self.root=tk.Tk()
        self.root.geometry("500x500")
        self.textbox1=tk.Entry(self.root)
        self.textbox2=tk.Entry(self.root)
        self.textbox3=tk.Entry(self.root)
        self.textbox4=tk.Entry(self.root)
        self.textbox5=tk.Entry(self.root)
        self.textbox6=tk.Entry(self.root)
        self.textbox7=tk.Entry(self.root)
        self.mylst=[]
        self.InitializeComponents()
        self.root.mainloop()
        self.Calculate()
        self.WriteToExcel()
    def InitializeComponents(self):
        tk.Label(self.root, text="Equation", font=('Comic Sans MS', 15)).grid(row=0)
        tk.Label(self.root,text="Velocity Approach Factor:").grid(row=1)
        self.textbox1.grid(row=1,column=1)
        tk.Label(self.root,text="Flow Area").grid(row=2)
        self.textbox2.grid(row=2,column=1)
        tk.Label(self.root,text="Inlet Fluid Density").grid(row=3)
        self.textbox3.grid(row=3,column=1)
        tk.Label(self.root,text="mass flow rate").grid(row=4)
        self.textbox4.grid(row=4,column=1)
        tk.Label(self.root,text="expansion factor").grid(row=5)
        self.textbox5.grid(row=5,column=1)
        tk.Label(self.root,text="Pressure Differential").grid(row=6)
        self.textbox6.grid(row=6,column=1)
        tk.Button(text='Calculate discharge Result',command=self.Calculate).grid(row=8,column=1,sticky=tk.W)
        
    def Calculate(self):
        velocityApproachFactor=float(self.textbox1.get())
        flowArea=float(self.textbox2.get())
        inletFluidDensity=float(self.textbox3.get())
        massFlowRate=float(self.textbox4.get())
        expansionFactor=float(self.textbox5.get())
        pressureDifferential=float(self.textbox6.get())
        Result=massFlowRate/(expansionFactor*velocityApproachFactor*flowArea)*math.sqrt(2*inletFluidDensity*pressureDifferential)
        self.mylst.append(Result)
        print(self.mylst)
        df=pd.DataFrame({'Data':self.mylst})
        writer=pd.ExcelWriter("C:\\Users\\mohamed ismail\\Desktop\\try.xlsx",engine='xlsxwriter')
        df.to_excel(writer,sheet_name='Sheet1')
        writer.save()
        
myObj=GUI()
myObj.InitializeComponents()