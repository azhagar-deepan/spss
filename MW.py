import SpssClient, spss
from rpstatspss import *


def MW(groups, code, testVar, precision, flag):
    SpssClient.StartClient()
    codeX = f"({code[0]},{code[1]})"
    median_decimal, Range_decimal, p_decimal = precision[0], precision[1], precision[2]
    if flag:
        tab = 1
        for eachGroup in groups:

            table = []
            table2 = []

            for var in testVar:
                table.append(f"{var}[MEDIAN, PTILE 25, PTILE 75] > {eachGroup}")
                table2.append(f"{var} [MEAN]")
            spss.Submit(
                f"""
                            CTABLES
                                /VLABELS VARIABLES={" ".join(testVar)} DISPLAY=LABEL
                                /TABLE {" + ".join(table2)} BY {eachGroup}
                                /CATEGORIES VARIABLES={eachGroup} ORDER=A KEY=VALUE EMPTY=EXCLUDE
                                /SLABELS VISIBLE=NO
                                /CRITERIA CILEVEL=95.
                            
            """
            )
            IndepT([eachGroup], code, testVar, precision, False)
            spss.Submit(
                f"""NPAR TESTS M-W={" ".join(testVar)} BY {eachGroup+codeX}.
                            CTABLES
                                /VLABELS VARIABLES={" ".join(testVar)} DISPLAY=LABEL
                                /TABLE {" + ".join(table)}
                                /CATEGORIES VARIABLES={eachGroup} ORDER=A KEY=VALUE EMPTY=EXCLUDE
                                /CRITERIA CILEVEL=95.                                                          
                            OUTPUT MODIFY
                                /SELECT ALL EXCEPT (TABLES)
                                /DELETEOBJECT DELETE = YES.            
                            OUTPUT MODIFY
                                /SELECT TABLES 
                                /IF COMMANDS=["T-Test"] SUBTYPES="Independent Samples Test" 
                                /DELETEOBJECT DELETE = YES.
                            OUTPUT MODIFY
                                /SELECT TABLES 
                                /IF COMMANDS=["NPar Tests"] 
                                SUBTYPES=["Mann Whitney Ranks"]
                                /DELETEOBJECT DELETE = YES.
            """
            )

            oDoc = SpssClient.GetDesignatedOutputDoc()
            oItems = oDoc.GetOutputItems()
            """
            for _ in range(oItems.Size()):
                oItem = oItems.GetItemAt(_)
                print(_,oItem.GetType())
            """
            for _ in range(oItems.Size()):
                # print("Current:",str(_))
                oItem = oItems.GetItemAt(_)
                if (
                    oItem is not None
                    and oItem.GetType() == SpssClient.OutputItemType.PIVOT
                ):
                    pTable = oItem.GetSpecificType()
                    if pTable.GetTitleText() == "Group Statistics":
                        pTable.SetUpdateScreen(False)
                        pTable.ShowAll()
                        Next_Table = accessObj(_ + 3)
                        C_Table = accessObj(_ - 2)
                        Prev_Table = accessObj(_ + 5)
                        Next_Table.SetUpdateScreen(False)
                        C_Table.SetUpdateScreen(False)
                        Next_Table.PivotManager().TransposeRowsWithColumns()
                        C_Table.ColumnLabelArray().InsertNewAfter(
                            1, 1, label="Statistical Significance"
                        )

                        sig = PivotData(Next_Table)[3]
                        # print(sig)

                        # print(Prev_Table.GetTitleText())
                        # print(1/0)
                        median = [
                            str(round(float(number), median_decimal))
                            for number in PivotData(Prev_Table)[0]
                        ]
                        Q1 = [
                            str(round(float(number), Range_decimal))
                            for number in PivotData(Prev_Table)[1]
                        ]
                        Q2 = [
                            str(round(float(number), Range_decimal))
                            for number in PivotData(Prev_Table)[2]
                        ]

                        Range = []
                        for i in range(len(Q1)):
                            Range.append(f"{Q1[i]},{Q2[i]}")

                        spss.Submit(
                            """
                                OUTPUT MODIFY
                                    /SELECT TABLES 
                                    /IF COMMANDS=["CTables(LAST)"] 
                                    SUBTYPES=["Custom Table"]
                                    /DELETEOBJECT DELETE = YES.
                        """
                        )
                        for ele in range(len(median)):
                            median[ele] = f"{median[ele]}({Range[ele]})"
                        pTable.ColumnLabelArray().InsertNewAfter(
                            1, 1, label="Statistical Significance"
                        )
                        pTable.ColumnLabelArray().InsertNewAfter(
                            1, 1, label="Median(Q1,Q3)"
                        )

                        C_Table.ClearSelection()
                        for j in range(2):
                            C_Table.ColumnLabelArray().SelectLabelAt(2, j)
                        C_Table.Group()
                        C_Table.ColumnLabelArray().SetValueAt(3, 1, "Median(Q1,Q3)")
                        ColPush(pTable, 1, 2, 4, sig, p_decimal)
                        ColPush(pTable, 0, 2, 2, median, False)
                        ColPush(C_Table, 0, 2, 2, sig, p_decimal)
                        for i in range(2):
                            pL = median[i::2]
                            ColPush(C_Table, 0, 1, i, pL, False)

                        spss.Submit(
                            """
                                            OUTPUT MODIFY 
                                                /SELECT TABLES 
                                                /IF COMMANDS=["T-Test"] SUBTYPES="Group Statistics"
                                                /TABLECELLS SELECT=["Std. Deviation"] APPLYTO=COLUMN HIDE=YES
                                                /TABLECELLS SELECT=["Std. Error Mean"] APPLYTO=COLUMN HIDE=YES
                                                /TABLECELLS SELECT=["Mean"] APPLYTO=COLUMN HIDE=YES.
                                            OUTPUT MODIFY
                                                /SELECT TABLES 
                                                /IF COMMANDS=["NPar Tests"] 
                                                SUBTYPES=["Mann Whitney Test Statistics"]
                                                /DELETEOBJECT DELETE = YES.
                        """
                        )
                        pTable.PivotManager().GetRowDimension(1).SetDimensionName(
                            "Variable"
                        )
                        pTable.PivotManager().GetColumnDimension(0).HideLabel()
                        pTable.HideTitle()
                        pTable.SetUpdateScreen(True)
                        Next_Table.SetUpdateScreen(True)
                        C_Table.SetUpdateScreen(True)
                        Title = "Output Table " + str(tab)
                        pTable.SetTitleText(Title)
                        tab += 1
                else:
                    # print("No Pivot")
                    pass

    else:
        for eachGroup in groups:
            table = []
            for var in testVar:
                table.append(f"{var}[MEDIAN, PTILE 25, PTILE 75] > {eachGroup}")
            # print(table)
            spss.Submit(
                f"""
                                    CTABLES
                                        /VLABELS VARIABLES={" ".join(testVar)} DISPLAY=LABEL
                                        /TABLE {" + ".join(table)}
                                        /CATEGORIES VARIABLES={eachGroup} ORDER=A KEY=VALUE EMPTY=EXCLUDE
                                        /CRITERIA CILEVEL=95.
                                    NPAR TESTS M-W={" ".join(testVar)} BY {eachGroup+codeX}.
            """
            )

    # print("Process Complete")
