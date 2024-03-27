import SpssClient, spss
from rpstatspss import *


def KW(groups, Range, testVar, precision,flag):
    SpssClient.StartClient()
    cat = list(range(Range[0], Range[1] + 1))
    mean_decimal, sd_decimal, p_decimal = precision[0], precision[1], precision[2]
    l = [str(i) for i in Range]
    codeX = "(" + " ".join(l) + ")"
    if flag:        
        tab = 1
        for eachGroup in groups:
            table = []
            table2 = []
            for var in testVar:
                table.append(f"{var} [MEDIAN, PTILE 25, PTILE 75]")
                table2.append(f"{var} [MEAN]")
            spss.Submit(
                f"""
                                CTABLES
                                    /VLABELS VARIABLES={" ".join(testVar)} DISPLAY=LABEL
                                    /TABLE {" + ".join(table2)} BY {eachGroup}
                                    /CATEGORIES VARIABLES={eachGroup} ORDER=A KEY=VALUE EMPTY=EXCLUDE
                                    /SLABELS VISIBLE=NO
                                    /CRITERIA CILEVEL=95.                                            
                                NPAR TESTS
                                    /K-W={" ".join(testVar)} BY {eachGroup+codeX}
                                    /MISSING ANALYSIS.
                                OUTPUT MODIFY
                                    /SELECT TABLES 
                                    /IF COMMANDS=["NPar Tests"] 
                                    SUBTYPES=["Kruskal Wallis Ranks"]
                                    /DELETEOBJECT DELETE = YES.         
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
                if pTable.GetTitleText() == "Test Statistics":
                    pTable.SetUpdateScreen(False)
                    pTable.PivotManager().TransposeRowsWithColumns()
                    sig = PivotData(pTable)[2]

                    C_Table = accessObj(_ - 3)
                    C_Table.ColumnLabelArray().InsertNewAfter(
                        1, 1, label="Statistical Significance"
                    )
                    ColPush(C_Table, 0, 2, len(cat), sig, p_decimal)

                    Next_Table = accessObj(_ + 2)

                    List = PivotData(Next_Table, Row=False)
                    # print(List)
                    median = []
                    Q1 = []
                    Q3 = []
                    for j in range(0, len(cat) * 3, 3):
                        for i in range(len(testVar)):
                            median.append(List[j][i])
                            Q1.append(List[j + 1][i])
                            Q3.append(List[j + 2][i])
                    # print(median,Q1,Q3)

                    median = [str(round(float(number), mean_decimal)) for number in median]

                    Q1 = [str(round(float(number), sd_decimal)) for number in Q1]
                    Q3 = [str(round(float(number), sd_decimal)) for number in Q3]

                    IQR = []
                    for i in range(len(Q1)):
                        IQR.append(f"{Q1[i]},{Q3[i]}")

                    for ele in range(len(median)):
                        median[ele] = f"{median[ele]}({IQR[ele]})"

                    spss.Submit(
                        """
                                    OUTPUT MODIFY
                                        /SELECT TABLES 
                                        /IF COMMANDS=["CTables(LAST)"] 
                                        SUBTYPES=["Custom Table"]
                                        /DELETEOBJECT DELETE = YES.
                                    OUTPUT MODIFY
                                        /SELECT TABLES 
                                        /IF COMMANDS=["NPar Tests(LAST)"] 
                                        SUBTYPES=["Kruskal Wallis Test Statistics"]
                                        /DELETEOBJECT DELETE = YES.
                            """
                    )

                    C_Table.ClearSelection()
                    for j in range(len(cat)):
                        C_Table.ColumnLabelArray().SelectLabelAt(2, j)
                    C_Table.Group()
                    C_Table.ColumnLabelArray().SetValueAt(3, 1, "Median(Q1,Q3)")

                    median_ar = []
                    for j in range(len(testVar)):
                        median_ar += median[j :: len(testVar)]
                    # print(median_ar)

                    for i in range(len(cat)):
                        pL = median_ar[i :: len(cat)]
                        ColPush(C_Table, 0, 1, i, pL, False)

                    C_Table.PivotManager().GetRowDimension(0).SetDimensionName("Variable")
                    C_Table.ShowAll()
                    Title = "Output Table " + str(tab)
                    Next_Table.SetTitleText(Title)
                    C_Table.SetTitleText(f"K-W {str(tab)}")
                    tab += 1
            else:
                pass
    else:
        for eachGroup in groups:
            table2 = []
            for var in testVar:
                table2.append(f"{var} [MEAN]")
            spss.Submit(
                f"""
                                NPAR TESTS
                                    /K-W={" ".join(testVar)} BY {eachGroup+codeX}
                                    /MISSING ANALYSIS.                               
            """
            )
