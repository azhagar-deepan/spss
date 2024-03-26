import SpssClient

SpssClient.StartClient()

def PivotData(TableObj,Row=False):
    datacells = TableObj.DataCellArray()
    PivotMgr = TableObj.PivotManager()
    
    RowDim = PivotMgr.GetRowDimension(0)
    ColDim = PivotMgr.GetColumnDimension(0)
    Data = []

    for i in range(RowDim.GetNumCategories()):
        temp = []
        for j in range(ColDim.GetNumCategories()):
            temp.append(datacells.GetUnformattedValueAt(i, j))
        Data.append(temp)
    if Row:
        return Data
    else:
        return list(zip(*Data))
    
def ValueExtract(TableObj, R, C):
    datacells = TableObj.DataCellArray()
    PivotMgr = TableObj.PivotManager()
    ColLabels = TableObj.ColumnLabelArray()
    RowLabels = TableObj.RowLabelArray()
    C -= 1
    Row = {}
    Col = {}
    Data = []

    RowDim = PivotMgr.GetRowDimension(0)
    ColDim = PivotMgr.GetColumnDimension(0)
    for i in range(RowDim.GetNumCategories()):
        temp = []
        for j in range(ColDim.GetNumCategories()):
            temp.append(datacells.GetUnformattedValueAt(i, j))
        Data.append(temp)

    for m in range(PivotMgr.GetNumColumnDimensions() - 1, -1, -1):
        ColDim = PivotMgr.GetColumnDimension(m)
        Col[ColDim.GetDimensionName()] = []
        for k in range(ColDim.GetNumCategories()):
            Col[ColDim.GetDimensionName()].append(ColDim.GetCategoryValueAt(k))

    for n in range(PivotMgr.GetNumRowDimensions() - 1, -1, -1):
        RowDim = PivotMgr.GetRowDimension(n)
        Row[RowDim.GetDimensionName()] = []
        for k in range(RowDim.GetNumCategories()):
            Row[RowDim.GetDimensionName()].append(RowDim.GetCategoryValueAt(k))
        unique_list = []
        [
            unique_list.append(item)
            for item in Row[RowDim.GetDimensionName()]
            if item not in unique_list
        ]
        Row[RowDim.GetDimensionName()] = unique_list
    S = Row[list(Row.keys())[-1]]
    e = R[0]
    chosen_variable = Row[list(Row.keys())[0]][e]
    chosen_choice = R[1]

    for i, variable in enumerate(Row[list(Row.keys())[0]]):
        if variable == chosen_variable:
            if len(list(Row.keys())) > 1:
                index = i * len(S) + chosen_choice - 1
            else:
                index = e
    return Data[index][C]

def ColData(TableObj,u,n,c):
    PivotMgr = TableObj.PivotManager()
    ColDim = PivotMgr.GetColumnDimension(0)
    RowDim = PivotMgr.GetRowDimension(1)
    Columndata = []    
    if u != 0:
        for i in range(RowDim.GetNumCategories()):
            Columndata.append(ValueExtract(TableObj,[i,u],c))
    else:
        for i in range(RowDim.GetNumCategories()):
            for j in range(1,n+1):
                Columndata.append(ValueExtract(TableObj,[i,j],c))
    return Columndata
    

def RowData(TableObj,n,u):
    PivotMgr = TableObj.PivotManager()
    RowDim = PivotMgr.GetRowDimension(1)
    ColDim = PivotMgr.GetColumnDimension(0)

    Rowdata = []
    
    for j in range(1,ColDim.GetNumCategories()+1):
        Rowdata.append(ValueExtract(TableObj,[n,u],j))
    return Rowdata


def ValuePush(TableObj, R, C,value,decimal):
    datacells = TableObj.DataCellArray()
    PivotMgr = TableObj.PivotManager()
    ColLabels = TableObj.ColumnLabelArray()
    RowLabels = TableObj.RowLabelArray()

    C -= 1
    Row = {}
    Col = {}
        
    for m in range(PivotMgr.GetNumColumnDimensions() - 1, -1, -1):
        ColDim = PivotMgr.GetColumnDimension(m)
        Col[ColDim.GetDimensionName()] = []
        for k in range(ColDim.GetNumCategories()):
            Col[ColDim.GetDimensionName()].append(ColDim.GetCategoryValueAt(k))

    for n in range(PivotMgr.GetNumRowDimensions() - 1, -1, -1):
        RowDim = PivotMgr.GetRowDimension(n)
        Row[RowDim.GetDimensionName()] = []
        for k in range(RowDim.GetNumCategories()):
            Row[RowDim.GetDimensionName()].append(RowDim.GetCategoryValueAt(k))
        unique_list = []
        [
            unique_list.append(item)
            for item in Row[RowDim.GetDimensionName()]
            if item not in unique_list
        ]
        Row[RowDim.GetDimensionName()] = unique_list
    S = Row[list(Row.keys())[-1]]
    e = R[0]
    chosen_variable = Row[list(Row.keys())[0]][e]
    chosen_choice = R[1]

    for i, variable in enumerate(Row[list(Row.keys())[0]]):
        if variable == chosen_variable:
            if len(list(Row.keys())) > 1:
                index = i * len(S) + chosen_choice - 1
            else:
                index = e
    if decimal != False:
        datacells.SetValueAt(index,C,value)
        datacells.SetNumericFormatAtWithDecimal(index,C,"#.#",decimal)    
    else:
        datacells.SetValueAt(index,C,value)

def RowPush(TableObj,n,u,List,decimal):
    PivotMgr = TableObj.PivotManager()
    RowDim = PivotMgr.GetRowDimension(1)
    ColDim = PivotMgr.GetColumnDimension(0)

    for j in range(1,ColDim.GetNumCategories()+1):
        ValuePush(TableObj,[n,u],j,List[j-1],decimal)

def ColPush(TableObj,u,n,c,List,decimal):
    PivotMgr = TableObj.PivotManager()
    RowDim = PivotMgr.GetRowDimension(1)
    ColDim = PivotMgr.GetColumnDimension(0)

    if u != 0:
        for i in range(RowDim.GetNumCategories()):
            ValuePush(TableObj,[i,u],c,List[i],decimal)
    else:
        f = RowDim.GetNumCategories()*len(list(range(1,n+1)))
        g = 0
        for i in range(RowDim.GetNumCategories()):
            for j in range(1,n+1):
                if g < f:
                    ValuePush(TableObj,[i,j],c,List[g],decimal)
                    g+=1
def accessObj(a):
    oItem = SpssClient.GetDesignatedOutputDoc().GetOutputItems().GetItemAt(a)
    if oItem.GetType() == SpssClient.OutputItemType.PIVOT:
        return oItem.GetSpecificType()
    else:
        print("Not a pivot")


SpssClient.StopClient()
