import SpssClient, spss
from rpstatspss import *


def IndepT(groups, code, testVar, precision, flag):
    SpssClient.StartClient()
    mean_decimal, sd_decimal, p_decimal = precision[0], precision[1], precision[2]
    code = f"({code[0]},{code[1]})"

    if flag:
        tab = 1
        for eachGroup in groups:
            table = []
            for var in testVar:
                table.append(f"{var} [MEAN]")

            spss.Submit(
                f"""
                                CTABLES
                                    /VLABELS VARIABLES={" ".join(testVar)} DISPLAY=LABEL
                                    /TABLE {" + ".join(table)} BY {eachGroup}
                                    /CATEGORIES VARIABLES={eachGroup} ORDER=A KEY=VALUE EMPTY=EXCLUDE
                                    /SLABELS VISIBLE=NO 
                                    /CRITERIA CILEVEL=95.
                                T-TEST GROUPS={eachGroup + code}
                                    /MISSING=ANALYSIS
                                    /VARIABLES={" ".join(testVar)}
                                    /ES DISPLAY(FALSE)
                                    /CRITERIA=CI(.95).
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
                if pTable.GetTitleText() == "Group Statistics":
                    pTable.SetUpdateScreen(False)
                    pTable.ShowAll()
                    Prev_Table = accessObj(_ - 2)
                    Next_Table = accessObj(_ + 1)
                    mean = [
                        str(round(float(number), mean_decimal))
                        for number in ColData(pTable, 0, 2, 2)
                    ]
                    sd = [
                        str(round(float(number), sd_decimal))
                        for number in ColData(pTable, 0, 2, 3)
                    ]
                    L1 = ColData(Next_Table, 1, 0, 2)
                    L2 = ColData(Next_Table, 0, 2, 5)
                    L = []
                    for val in L1:
                        if float(val) > 0.05:
                            L.append(L2[0])
                        else:
                            L.append(L2[1])
                        L2 = L2[2:]

                    for ele in range(len(mean)):
                        mean[ele] = mean[ele] + "±" + sd[ele]
                    pTable.ColumnLabelArray().InsertNewAfter(
                        1, 1, label="Statistical Significance"
                    )
                    pTable.ColumnLabelArray().InsertNewAfter(1, 1, label="Mean±S.D")
                    ColPush(pTable, 1, 2, 4, L, p_decimal)
                    ColPush(pTable, 0, 2, 2, mean, False)
                    Prev_Table.ColumnLabelArray().InsertNewAfter(
                        1, 1, label="Statistical Significance"
                    )
                    Prev_Table.ClearSelection()
                    for j in range(2):
                        Prev_Table.ColumnLabelArray().SelectLabelAt(2, j)
                    Prev_Table.Group()
                    Prev_Table.ColumnLabelArray().SetValueAt(3, 1, "Mean±S.D")
                    for i in range(2):
                        pL = mean[i::2]
                        ColPush(Prev_Table, 0, 1, i, pL, False)

                    ColPush(Prev_Table, 0, 2, 2, L, p_decimal)
                    Prev_Table.PivotManager().GetRowDimension(0).SetDimensionName(
                        "Variable"
                    )

                    Prev_Table.PivotManager().GetColumnDimension(1).HideLabel()
                    Prev_Table.Autofit()  # Last added line <<<<<<<<<<<<<<<<<<<<<<
                    spss.Submit(
                        """                       
        OUTPUT MODIFY 
        /SELECT TABLES 
        /IF COMMANDS=["T-Test"] SUBTYPES="Group Statistics"
        /TABLECELLS SELECT=["Std. Deviation"] APPLYTO=COLUMN HIDE=YES
        /TABLECELLS SELECT=["Std. Error Mean"] APPLYTO=COLUMN HIDE=YES
        /TABLECELLS SELECT=["Mean"] APPLYTO=COLUMN HIDE=YES.
        """
                    )
                    pTable.PivotManager().GetColumnDimension(0).HideLabel()
                    pTable.PivotManager().GetRowDimension(1).SetDimensionName(
                        "Variable"
                    )

                    # print(pTable.PivotManager(). GetNumColumnDimensions())

                    pTable.SetUpdateScreen(True)
                    Title = "Output Table " + str(tab)
                    pTable.SetTitleText(Title)
                    pTable.HideTitle()
                    tab += 1
            else:
                pass

        spss.Submit(
            """
        OUTPUT MODIFY
        /SELECT TABLES 
        /IF COMMANDS=["T-Test"] SUBTYPES="Independent Samples Test" 
        /DELETEOBJECT DELETE = YES.
        """
        )
    else:
        for eachGroup in groups:
            spss.Submit(
                f"""
            T-TEST GROUPS={eachGroup + code}
                  /MISSING=ANALYSIS
                  /VARIABLES={" ".join(testVar)}
                  /ES DISPLAY(FALSE)
                  /CRITERIA=CI(.95).
  """
            )
