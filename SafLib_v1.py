'''
Methods List :
HelloWorld

'''

from enum import unique
from importlib import invalidate_caches
from operator import mod
#from turtle import width
from scriptforge import CreateScriptService

## Global Constants
DEFAULT_INV_LENGTH=5486 # 18 Feet
DISCARD_LENGTH=50


#Cut List Column Name with Index
CutListSheet="CutList"
MaterialCategory=0
MaterialToCut=1
LengthToCut=3
QtyToCut=4
################################
#Inventory List Column Names with Index(s)
InventorySheet="InventoryList"
Inv_Material=1
Inv_Length=2
Inv_Qty=3
################################

#Writed Hello World in Cell A1
def HelloWorld(args=None):
    doc=CreateScriptService("Calc")
    #doc.setValue("A1","Hello World !!!")
    print("Mod of 3500 to 1500 is ",3500//750)

def MaterialWiseQuantity(args=None):
    doc=CreateScriptService("Calc")
    CutList=doc.GetValue("~.A1:E12")
    InventoryList=doc.GetValue("~.A15:D20")
    print(doc.Height("A2:E12"))
    for i in range(doc.Height("A1:E12")):
        print(i," - ", CutList[i][2])
        for j in range(doc.Height("A15:D20")):
            if CutList[i][1]==InventoryList[j][1]:
                print(InventoryList[j])

def UniqueMaterialList(args=None):
    doc=CreateScriptService("Calc")
    unique_list=[]
    CutList=doc.getValue("B1:B12")
    for i in CutList:
        if i not in unique_list:
            unique_list.append(i)
    for i in unique_list:
        print(i)

def NestMaterial(args=None):
    print("*************** Starting new Operation ******************")
    
    doc=CreateScriptService("Calc")
    
    doc.sortRange(doc.currentselection,(1,2,4),("Asc","Asc","Desc"),casesensitive=True)
    Sort_InventoryRange()
    #doc.sortRange("InventoryList.B2:D25",(1,2,3),("Asc","Asc","Asc"),casesensitive=True)
    
    #CutList=doc.getValue("CutList.A1:E1") # first nest the first record properly, then will loop the rest
    CutList=doc.getValue(doc.currentselection)
    print("#### Total Cut list is :",CutList)
    
    for c_cutlist in CutList:

        Qty_to_cut=int(c_cutlist[4])

        print("##### Cut List to Cut : ",c_cutlist, " Qty to cut is ",Qty_to_cut)
        
        while Qty_to_cut>0:
            print("#### A] Qty to cut is ",Qty_to_cut)
            print("Calling Inventory Issue to get record to use for nesting...")
            s_record = IssueMaterialfromInventory(c_cutlist)
            print("##### Inventory record to use : ", s_record)
            #Nest_Material_from__Record(s_record,CutList)
            print("#### Value of s_record[1] ", s_record[1]," and CutList[3] is ", c_cutlist[3])
            Qty_cut=int(s_record[1])//int(c_cutlist[3])
            print("#### Qty_cut is ",Qty_cut)
            if(Qty_to_cut>Qty_cut):
                Nest_Material_from__Record(Qty_cut,s_record,c_cutlist)
            else:
                Nest_Material_from__Record(Qty_to_cut,s_record,c_cutlist)
            if(Qty_cut!=0):
                Qty_to_cut=Qty_to_cut-Qty_cut
            else:
                Qty_to_cut=Qty_to_cut-1
            print("#### B] Qty to cut is ",Qty_to_cut)
            
            #Nest_Material_from__Record(s_record,CutList)
        

def IssueMaterialfromInventory(s_MaterialList =[]):
    print("@@@@ Inside Issue Material from Inventory with Material :",s_MaterialList)
    doc=CreateScriptService("Calc")
    InventoryList=doc.getValue("InventoryList.B2:D25")
    inv_counter=0
    
    for i in InventoryList:
        print("@@@@ Inventory record for matching : ",i)
        
        ### Match Material and Length required to Nest
        
        if(i[0]==s_MaterialList[1] and i[1]>=s_MaterialList[3]):
            print("Inventory matched, length available, i[1] is",i[1]," and s_materialList[2] is ", s_MaterialList[3])
            cell=doc.offset("InventoryList.B2",inv_counter,2)
            print("@@@@ Cell value is : ",cell)
            '''
            if Qty is already at 1, then clear the value since, last item is issued.
            '''
            if(i[2]>1):
                #if Value is greater than 1, the deduct 1 Nos and return record
                print("@@@@ value updated in Inventory List at cell : ",cell," is ",i[2]-1)
                doc.SetValue(cell,i[2]-1)
            elif(i[2]==1):
                # if value is 1, create one default entry if 
                print("@@@@ In ELIF, Qty to Handle is 1")
                print("@@@@ In ELIF Qty available in inventory is : ", i[2])    
                cell=doc.offset("InventoryList.A1",inv_counter+1,0,0,4)
                print("@@@@ In ELIF Value of Cell to clear is : ",cell)
                doc.clearvalues(cell)
                Sort_InventoryRange()
                Check_Default_Length_Exists(s_MaterialList[1],True)
            else:
                # 
                print("@@@@ in ELSE Qty available in inventory is : ", i[2])    
                cell=doc.offset("InventoryList.A1",inv_counter+1,0,0,4)
                print("@@@@ in ELSE  Value of Cell to clear is : ",cell)
                doc.clearvalues(cell)
                Sort_InventoryRange()
            return i
        inv_counter=inv_counter+1
    # if reached here then add default value record and return that.
    ### Test if record does not exists in inventory, what happens
    print("@@@@ Reached end of inventory List. Record Not found. Now adding defauly entry for record and returning that")
    arr_data=((5),(s_MaterialList[1]),(DEFAULT_INV_LENGTH),(0))
    cell=doc.offset("InventoryList.A1",doc.LastRow("InventoryList"),0,0,4)
    print("^^^^ range_cell is ",cell)
    print("^^^^ New Record inserted in Inventory is : ", arr_data)
    doc.setValue(cell,arr_data)
    print("@@@@ last row is ", doc.LastRow("InventoryList"))
    cell=doc.offset("InventoryList.B1",doc.LastRow("InventoryList")-1,0,0,3) 
    print("@@@@ Value of cell is ",cell)
    return doc.getValue(cell)
    Sort_InventoryRange()
    #Re-run the loop to get the record and return that, if return fails.

def Nest_Material_from__Record(QtyToNest,s_inv_record=[],s_material=[]):
    
    
    print("$$$$ Inside Nest Material. s_record is ",s_inv_record)
    print("$$$$ Material to nest is ",s_material)
    
    
    doc=CreateScriptService("Calc")

    #No_items_to_nest=int(s_material[4])
    No_items_to_nest=int(QtyToNest)
    length_to_Nest=int(s_material[3])
    length_from_inventory=int(s_inv_record[1])
    qty_from_inventory=int(s_inv_record[2])

    print("$$$$ Number of items to nest here is ",No_items_to_nest)
    
    while No_items_to_nest > 0:
        length_from_inventory=length_from_inventory-length_to_Nest
        No_items_to_nest=No_items_to_nest-1

    print("$$$$ Last Row used is : ",doc.LastRow("InventoryList"))
    print("$$$$ No of items remaining to nest : ", No_items_to_nest)
    print("$$$$ length remaining in inventory is :", length_from_inventory)
    print("$$$$ Qty from Inventory is :",qty_from_inventory)
    AddRecord=True
    if(qty_from_inventory>0):
        if(length_from_inventory>DISCARD_LENGTH):
            arr_data=((1),(s_material[1]),(length_from_inventory),(1))
        else:
            Check_Default_Length_Exists(s_material[1],True)
            AddRecord=False
            #arr_data=((2),(s_material[1]),(DEFAULT_INV_LENGTH),(0))
    else: #if qty is Less than or equals zero, then set negetive value. This denotes amount of angles to order       
        if(qty_from_inventory==0): # if alredy at 0, then set it as -1.
            qty_from_inventory=-1
        if(length_from_inventory>DISCARD_LENGTH):
            Check_Default_Length_Exists(s_material[1],True)
            arr_data=((3),(s_material[1]),(length_from_inventory),qty_from_inventory)
        else:
            # If Qty is Negetive and Length remaining is to be Discarded, then
            # New Entry should have length as 0 since no more nesting of this is poss
            #arr_data=((4),(s_material[1]),(DEFAULT_INV_LENGTH),qty_from_inventory-1)
            #arr_data=((4),(s_material[1]),(DEFAULT_INV_LENGTH),qty_from_inventory)
            arr_data=((4),(s_material[1]),(0),qty_from_inventory)
    

    # if Default Length entry already exists
    if(AddRecord==True):
        cell=doc.offset("InventoryList.A1",doc.LastRow("InventoryList"),0,0,4)
        print("$$$$ range_cell is ",cell)
        print("$$$$ New Record inserted in Inventory is : ", arr_data)
        doc.setValue(cell,arr_data)    
        Sort_InventoryRange()
        #doc.sortRange("InventoryList.B2:D25",(1,2,3),("Asc","Asc","Asc"),casesensitive=True)

def Check_Default_Length_Exists(str_material,AddEntry):
    print("^^^^ in default length check")
    doc=CreateScriptService("Calc")
    InventoryList=doc.getValue("InventoryList.B2:D25")
    for i in InventoryList:
        if(i[0]==str_material and i[1]==DEFAULT_INV_LENGTH and i[2]==0):
            print("^^^^ Default length already exists")
            return True
            break
    
    # Add Default Length since it does not exist
    if(AddEntry==True):

        arr_data=((2),(str_material),(DEFAULT_INV_LENGTH),(0))
        
        cell=doc.offset("InventoryList.A1",doc.LastRow("InventoryList"),0,0,4)
        print("^^^^ range_cell is ",cell)
        print("^^^^ New Record inserted in Inventory is : ", arr_data)
        doc.setValue(cell,arr_data)    
        Sort_InventoryRange()
        return True # Default Length is added, now return True
    return False

def Sort_InventoryRange(args=None):
    print("--- In Sort Inventory Funtion ---")
    doc=CreateScriptService("Calc")
    doc.sortRange("InventoryList.B2:D35",(1,2,3),("Asc","Asc","Desc"),casesensitive=True)

g_exportedScripts = (HelloWorld, MaterialWiseQuantity,UniqueMaterialList, NestMaterial,IssueMaterialfromInventory,)

