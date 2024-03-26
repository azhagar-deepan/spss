import SpssClient, spss
from rpstatspss import *

def Run(groups, code, testVar):
    testVar = " ".join(testVar)
    code = f"({code[0]},{code[1]})"
    for eachGroup in groups:
        spss.Submit(f"""T-TEST GROUPS={eachGroup + code}
  /MISSING=ANALYSIS
  /VARIABLES={testVar}
  /ES DISPLAY(FALSE)
  /CRITERIA=CI(.95).""")

def Format(mean_decimal,sd_decimal,p_decimal):
    SpssClient.StartClient()
    oDoc = SpssClient.GetDesignatedOutputDoc()
    oItems = oDoc.GetOutputItems() 
    
    tab = 1   
    for _ in range(oItems.Size()):
        #print(_)
        oItem = oItems.GetItemAt(_)        
        if oItem is not None and oItem.GetType() == SpssClient.OutputItemType.PIVOT: 
    
            pTable = oItem.GetSpecificType()
            if pTable.GetTitleText() == "Group Statistics":
                pTable.SetUpdateScreen(False)
                pTable.ShowAll()
                Next_Table = accessObj(_+1)
                mean = [str(round(float(number), mean_decimal)) for number in ColData(pTable,0,2,2)]
                sd = [str(round(float(number), sd_decimal)) for number in ColData(pTable,0,2,3)] 
                L1 = ColData(Next_Table,1,0,2)
                #L = [str(round(float(number), 3)) for number in L]
                L2 = ColData(Next_Table,0,2,5)
                L = []
                for val in L1 :
                    if float(val) > 0.05:
                        L.append(L2[0])  
                    else:
                        L.append(L2[1])  
                    L2 = L2[2:]
                #mean = ColData(pTable,0,2,2)
                #sd = ColData(pTable,0,2,3)
    
                for ele in range(len(mean)):
                    mean[ele] = mean[ele] + "±" + sd[ele]
                #print(mean)
                pTable.ColumnLabelArray().InsertNewAfter(1,1,label="Statistical Significance")
                pTable.ColumnLabelArray().InsertNewAfter(1,1,label="Mean±S.D")
                ColPush(pTable,1,2,4,L,p_decimal)
                ColPush(pTable,0,2,3,mean,False)                
                spss.Submit('''
    OUTPUT MODIFY 
      /SELECT TABLES 
      /IF COMMANDS=["T-Test"] SUBTYPES="Group Statistics"
       /TABLECELLS SELECT=["Std. Deviation"] APPLYTO=COLUMN HIDE=YES
       /TABLECELLS SELECT=["Std. Error Mean"] APPLYTO=COLUMN HIDE=YES
       /TABLECELLS SELECT=["Mean"] APPLYTO=COLUMN HIDE=YES.
      ''')
                pTable.PivotManager().GetRowDimension(1).SetDimensionName("Variable")
                pTable.PivotManager().GetColumnDimension(0).HideLabel()
                pTable.SetUpdateScreen(True)
                Title = "Output Table "+str(tab)
                pTable.SetTitleText(Title)
                pTable.HideTitle()
                tab += 1
        else:
            pass           
        
    spss.Submit('''
    OUTPUT MODIFY
      /SELECT TABLES 
      /IF COMMANDS=["T-Test"] SUBTYPES="Independent Samples Test" 
      /DELETEOBJECT DELETE = YES.
      ''')