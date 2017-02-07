from openpyxl import load_workbook
from tkinter import * 
from collections import defaultdict
import pickle 
import glob



unique = pickle.load( open( "unique.p", "rb" ) )
listofcells = ["C8", "D8", "C9","D9","C21"]
labeldicc =  {}


stringvardicc = {}

currentcell = 0


def gradeCell():


    global currentcell
    fenetre = Tk()
    cell = listofcells[currentcell]
    currentrow = 0
    labeldicc[0,cell,"cell"] = Label(fenetre, text=cell).grid(row=currentrow,column=1,pady=20)
    currentrow = currentrow + 1
    for elementindex, element in enumerate(unique[cell]):
            labeldicc[1,cell,elementindex,"formula"] = Label(fenetre, text=unique[cell][elementindex][0]).grid(row=currentrow, padx=10,column=0)
            labeldicc[2,cell,elementindex,"formula"] = Label(fenetre, text=unique[cell][elementindex][1]).grid(row=currentrow,padx=10,column=1)
            stringvardicc[cell,elementindex] = StringVar()
            stringvardicc[cell,elementindex].set(unique[cell,elementindex])
            labeldicc[3,cell,elementindex,"formula"] = Entry(fenetre, textvariable=stringvardicc[cell,elementindex], width=30).grid(row=currentrow,column=2)
            currentrow = currentrow + 1
    def nextcell():
        on_closing() 
        global currentcell
        currentcell = currentcell + 1
        gradeCell() 
    b = Button(fenetre, text="nextcell", command=nextcell).grid(column=1,pady=25)  
      
    def on_closing():
        savegrades()
        fenetre.destroy()  
    def savegrades():
        for elementindex, element in enumerate(unique[cell]):
             unique[cell,elementindex]=stringvardicc[cell,elementindex].get()
        pickle.dump( unique, open( "unique.p", "wb" ) )

    fenetre.configure(background='white')    
    fenetre.protocol("WM_DELETE_WINDOW", on_closing)
    w, h = fenetre.winfo_screenwidth(), fenetre.winfo_screenheight()
    fenetre.geometry("%dx%d+0+0" % (w, h))
    fenetre.mainloop()



gradeCell()