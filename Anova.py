import SpssClient, spss
from rpstatspss import *


def Anova(groups, cat, testVar, precision, flag):
    SpssClient.StartClient()
    mean_decimal, sd_decimal, p_decimal = precision[0], precision[1], precision[2]
    tab = 1
    for eachGroup in groups:
        table = []
        table2 = []
        for var in testVar:
            table.append(f"{var} [MEAN,STDDEV]")
            table2.append(f"{var} [MEAN]")
        spss.Submit(
            f"""
                            CTABLES
                                /VLABELS VARIABLES={" ".join(testVar)} DISPLAY=LABEL
                                /TABLE {" + ".join(table2)} BY {eachGroup}
                                /CATEGORIES VARIABLES={eachGroup} ORDER=A KEY=VALUE EMPTY=EXCLUDE
                                /SLABELS VISIBLE=NO
                                /CRITERIA CILEVEL=95.
                            ONEWAY {" ".join(testVar)} BY {eachGroup}
                                /MISSING ANALYSIS
                                /CRITERIA=CILEVEL(0.95).
                            CTABLES
                                /VLABELS VARIABLES={" ".join(testVar)} DISPLAY=LABEL
                                /TABLE {" + ".join(table)} BY {eachGroup}
                                /CATEGORIES VARIABLES={eachGroup} ORDER=A KEY=VALUE EMPTY=EXCLUDE
                                /SLABELS VISIBLE=YES
                                /CRITERIA CILEVEL=95.                                                           
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
            if pTable.GetTitleText() == "ANOVA":
                pTable.SetUpdateScreen(False)
                pTable.ShowAll()
                Prev_Table = accessObj(_ + 2)

                List = PivotData(Prev_Table, Row=False)

                mean = []
                sd = []
                for j in range(0, cat * 2, 2):
                    mean.append(List[j])
                    sd.append(List[j + 1])

                mean = [
                    str(round(float(number), mean_decimal))
                    for cap in mean
                    for number in list(cap)
                ]
                sd = [
                    str(round(float(number), sd_decimal))
                    for cap in sd
                    for number in list(cap)
                ]
                # print(mean,sd)
                for ele in range(len(mean)):
                    mean[ele] = mean[ele] + "±" + sd[ele]
                # print(mean)
                if flag:
                    spss.Submit(
                        """
                                                OUTPUT MODIFY
                                                    /SELECT TABLES 
                                                    /IF COMMANDS=["Oneway(LAST)"] 
                                                    SUBTYPES=["ANOVA"]
                                                    /DELETEOBJECT DELETE = YES.
                                                OUTPUT MODIFY
                                                    /SELECT TABLES 
                                                    /IF COMMANDS=["CTables(LAST)"] 
                                                    SUBTYPES=["Custom Table"]
                                                    /DELETEOBJECT DELETE = YES.
                            """
                    )

                Next_Table = accessObj(_ - 2)
                Next_Table.ColumnLabelArray().InsertNewAfter(
                    1, 1, label="Statistical Significance"
                )
                Next_Table.ClearSelection()
                for j in range(cat):
                    Next_Table.ColumnLabelArray().SelectLabelAt(2, j)
                Next_Table.Group()

                Next_Table.ColumnLabelArray().SetValueAt(3, 1, "Mean±S.D")

                mean_ar = []
                for j in range(len(testVar)):
                    mean_ar += mean[j :: len(testVar)]
                # print(mean_ar)
                for i in range(cat):
                    pL = mean_ar[i::cat]
                    ColPush(Next_Table, 0, 1, i, pL, False)
                Next_Table.Autofit()
                ColPush(
                    Next_Table, 0, 2, cat, ColData(pTable, 1, 0, 5), p_decimal
                )  # Take a note here

                Next_Table.PivotManager().GetRowDimension(0).SetDimensionName(
                    "Variable"
                )

                Title = "Output Table " + str(tab)
                Next_Table.SetTitleText(Title)
                Next_Table.HideTitle()
                tab += 1
        else:
            pass
