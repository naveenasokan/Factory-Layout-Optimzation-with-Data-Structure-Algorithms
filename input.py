import pandas as pd
import numpy
import os

ROOT = os.getcwd()

#Sample function preparation
def Creatsamplefunction(df_comp, df_part, WORKCENTRE):

    df_part["Concat"] = df_part["Material Sales Model"].astype(
        str) + df_part["Material"].astype(str) + df_part["Workcenter (Routing)"].astype(str)

    df_comp = df_comp.loc[df_comp["Work Center"].isin(WORKCENTRE)]
    df_comp = df_comp.loc[:, ["Serial Number",
                              "Sales Model", "Work Center", "Mech Class", "Item ID"]]
    df_comp["Concat"] = df_comp["Sales Model"].astype(
        str) + df_comp["Item ID"].astype(str) + df_comp["Work Center"].astype(str)

    df_cons = pd.merge(df_comp, df_part[[
                       "SupplyArea", "Storage Bin", "Concat"]], how='left', on=['Concat'])
    df_cons.dropna(subset=["Storage Bin"], inplace=True)
    df_cons.to_csv(os.path.join(os.getcwd(), "Input",
                   "samplefunction.csv"), index=False)


# input data preparation
def CreateInputSheet(df1,SOURCE_TEXT_ARRAY,WORKCENTRE):
    df1.dropna(subset=["Storage Bin"], inplace=True)
    df_consolidated = pd.DataFrame()
    print("Creating Input..... Please wait")
    
    for models in SOURCE_TEXT_ARRAY:
        df_temp = df1.loc[(
            df1["Sales Model"] == models) & (df1["Work Center"] == WORKCENTRE[0])]
        if not df_temp.shape[0] == 0:
            temp_serialno = df_temp["Serial Number"].value_counts(
            ).index[0]
            df_temp = df_temp.loc[df_temp["Serial Number"]
                                    == temp_serialno]

            df_consolidated = df_consolidated.append(df_temp.loc[:, [
                                                        "Sales Model", "Mech Class", "Serial Number", "Item ID", "SupplyArea", "Storage Bin"]], ignore_index=True)

    df_consolidated["Sales Model"] = WORKCENTRE[0].replace(
        "16WS", "") + "_" + df_consolidated["Sales Model"].astype(str)
    df_consolidated.drop_duplicates(
        subset=["Sales Model", "Mech Class", "Serial Number", "Item ID", "Storage Bin"], inplace=True)
    df_consolidated.to_excel(os.path.join(ROOT, "sample.xlsx"), index=False)

    df_consolidated.to_excel(os.path.join(
        ROOT, "Input\Layout Input", f"{WORKCENTRE[0]}_sample.xlsx"), index=False)


if __name__ == "__main__":
    pass
