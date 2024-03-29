import SpssClient, spss
from rpstatspss import *


def PairT(testVar1, testVar2, precision, flag):
    SpssClient.StartClient()
    mean_decimal, sd_decimal, p_decimal = precision[0], precision[1], precision[2]
    if flag:
        spss.Submit(
            f"""
                            T-TEST PAIRS={" ".join(testVar1)} WITH {" ".join(testVar2)}(PAIRED)
                                  /ES DISPLAY(FALSE)
                                  /CRITERIA=CI(.9500)
                                  /MISSING=ANALYSIS.
                            OUTPUT MODIFY
                                /SELECT ALL EXCEPT (TABLES)
                                /DELETEOBJECT DELETE = YES.            
        """
        )
    oDoc = SpssClient.GetDesignatedOutputDoc()
    oItems = oDoc.GetOutputItems()
    for _ in range(oItems.Size()):
        oItem = oItems.GetItemAt(_)
        if oItem is not None and oItem.GetType() == SpssClient.OutputItemType.PIVOT:
            pTable = oItem.GetSpecificType()
            if pTable.GetTitleText() == "Paired Samples Statistics":
                pTable.SetUpdateScreen(True)
                Next_Table = accessObj(_ + 2)
                mean = [
                    str(round(float(number), mean_decimal))
                    for number in PivotData(pTable)[0]
                ]
                sd = [
                    str(round(float(number), sd_decimal))
                    for number in PivotData(pTable)[2]
                ]

                for ele in range(len(mean)):
                    mean[ele] = mean[ele] + "±" + sd[ele]
                # print(mean,sd)
                pTable.ColumnLabelArray().InsertNewAfter(
                    1, 3, label="Statistical Significance"
                )

                sig = list(PivotData(Next_Table)[7])

                # print(sig)
                for i in range(1, len(sig) * 2 + 1, 2):
                    sig.insert(i, "")

                ColPush(pTable, 0, 2, 4, sig, False)
                ColPush(pTable, 0, 2, 2, mean, False)
                for i in range(len(sig)):
                    pTable.DataCellArray().SetNumericFormatAtWithDecimal(
                        i, 4, "#.#", p_decimal
                    )
                spss.Submit(
                    """                       
                            OUTPUT MODIFY 
                                /SELECT TABLES 
                                /IF COMMANDS=["T-Test"] SUBTYPES="Paired Samples Statistics"
                                /TABLECELLS SELECT=["Std. Error Mean"] APPLYTO=COLUMN HIDE=YES
                                /TABLECELLS SELECT=["Mean"] APPLYTO=COLUMN HIDE=YES.
                            OUTPUT MODIFY
                                /SELECT TABLES 
                                /IF COMMANDS=["T-Test"] SUBTYPES=["Paired Samples Correlations","Paired Samples Test"] 
                                /DELETEOBJECT DELETE = YES.                
                       """
                )
                pTable.ColumnLabelArray().SetValueAt(1, 2, "Mean±S.D")
                pTable.SetUpdateScreen(True)
                pTable.SetTitleText("Output Table")
