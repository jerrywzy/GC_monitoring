import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import glob

## dCt of PCR replicates 
dct = {}
final_dct = pd.DataFrame()

for eachfile in glob.glob("*/*.xlsm"):
    df = pd.read_excel(eachfile, sheet_name="Analysis")
    
    
    # find position of cell QR and Samples
    rows, cols = np.where(df == 'dCt of PCR replicates')
    
    # subset to only required section and clean up
    df2 = df.iloc[(int(rows)):(int(rows)+16), (int(cols)):(int(cols)+13)].reset_index(drop=True)
    # reset header and body
    df2.columns = df2.iloc[0]
    df3 = df2[1:]
    df4 = df3[df3['dCt of PCR replicates'] != 'NFW']   # full table
    
    # find position of cell for NC
    rows, cols = np.where(df == 'Min of PCR replicates')
    df_min = df.iloc[(int(rows)):(int(rows)+1), (int(cols)):(int(cols)+15)].reset_index(drop=True)
    df_min_2 = df_min.iloc[: , 2:]
    df_min_2.columns = df4.columns

    df5 = pd.concat([df4, df_min_2]).sort_index()
    
    # run name for x axis
    # print(eachfile)
    run_name = eachfile.split("/")[1].split("_")[0] + "_" + eachfile.split("/")[1].split("_")[1]
    df5['Run'] = run_name
    
    # date 
    raw_date = str(run_name.split('_')[0])
    day = raw_date[0] + raw_date[1]
    month = raw_date[2] + raw_date[3]
    year = raw_date[4] + raw_date[5] + raw_date[6] + raw_date[7]
    df5['RawDate'] = str(year + "/" + month + "/" + day)
    df5['Date'] = pd.to_datetime(df5.RawDate)
    
    # operator and kit
    df = pd.read_excel(eachfile, sheet_name="Report")
    rows, cols = np.where(df == 'Analysed By/ Sign :')
    df2 = df.iloc[(int(rows)):(int(rows)+1), (int(cols)):(int(cols)+2)].reset_index(drop=True)
    operator = df2.iloc[0][1]
    
    rows, cols = np.where(df == 'GASTROClyr Lot Number:')
    df2 = df.iloc[(int(rows)):(int(rows)+1), (int(cols)):(int(cols)+2)].reset_index(drop=True)
    lot = df2.iloc[0][1]
    
    df5['Operator'] = operator
    df5['KitLot'] = lot
    
    # final dataframe
    final_dct = pd.concat([final_dct, df5])
    final_dct = final_dct[['Date', 'Run', 'KitLot', 'Operator', 'dCt of PCR replicates', 'G2', 'G1', 'G3', 'G4', 'G5', 
                            'G6', 'G7', 'G8', 'G9', 'G10', 'G11', 'G12', ]]
    
final_dct["Date"] = pd.to_datetime(final_dct["Date"])
final_dct = final_dct.sort_values(by="Date")
final_dct.to_csv('GCMonitoring/dct.csv', index=False)
    

### NC
nc = {}
final_nc = pd.DataFrame()

for eachfile in glob.glob("*/*.xlsm"):
    df = pd.read_excel(eachfile, sheet_name="Analysis")
    
    # find position of cell
    rows, cols = np.where(df == 'Average of PCR replicates')
    
    # subset to only required section and clean up
    df2 = df.iloc[(int(rows)):(int(rows)+17), (int(cols)):(int(cols)+14)].reset_index(drop=True)
    # reset header and body
    df2.columns = df2.iloc[0]
    df2_2 = df2.rename(columns={np.nan: 'Average of PCR replicates'})
    df3 = df2_2[1:]   # one row down
    df3_2 = df3.iloc[0:, 1:]   # start from second column
    df3_3 = df3_2.fillna(40)
    
    # run name for x axis
    run_name = eachfile.split("/")[1].split("_")[0] + "_" + eachfile.split("/")[1].split("_")[1]
    df3_3['Run'] = run_name
    
    # date 
    raw_date = str(run_name.split('_')[0])
    day = raw_date[0] + raw_date[1]
    month = raw_date[2] + raw_date[3]
    year = raw_date[4] + raw_date[5] + raw_date[6] + raw_date[7]
    df3_3['RawDate'] = str(year + "/" + month + "/" + day)
    df3_3['Date'] = pd.to_datetime(df3_3.RawDate)
    
    # operator and kit
    df = pd.read_excel(eachfile, sheet_name="Report")
    rows, cols = np.where(df == 'Analysed By/ Sign :')
    df2 = df.iloc[(int(rows)):(int(rows)+1), (int(cols)):(int(cols)+2)].reset_index(drop=True)
    operator = df2.iloc[0][1]
    
    rows, cols = np.where(df == 'GASTROClyr Lot Number:')
    df2 = df.iloc[(int(rows)):(int(rows)+1), (int(cols)):(int(cols)+2)].reset_index(drop=True)
    lot = df2.iloc[0][1]
    
    df3_3['Operator'] = operator
    df3_3['KitLot'] = lot
    
    # final dataframe
    final_nc = pd.concat([final_nc, df3_3])
    final_nc = final_nc[['Date', 'Run', 'KitLot', 'Operator', 'Average of PCR replicates', 
                          'G2', 'G1', 'G3', 'G4', 'G5', 'G6', 'G7', 'G8', 'G9', 'G10', 'G11', 'G12']]

final_nc["Date"] = pd.to_datetime(final_nc["Date"])
final_nc = final_nc.sort_values(by="Date")
final_nc.to_csv('GCMonitoring/nc.csv', index=False)

### Raw Score (QR only)
raw_score = {}
final_raw = pd.DataFrame()

for eachfile in glob.glob("*/*.xlsm"):
    df = pd.read_excel(eachfile, sheet_name="Analysis")
    
    # find position of cell
    rows, cols = np.where(df == 'Raw score')
    
    # subset to only required section and clean up
    df2 = df.iloc[(int(rows)):(int(rows)+4), (int(cols)):(int(cols)+3)].reset_index(drop=True)
    # reset header and body
    df3 = df2.iloc[[2,3], 1:]   # take second and third row and start from second column
    df4 = df3.replace(['QR 1'],'QR1')
    df5 = df4.replace(['QR 2'],'QR2')
    df5.columns = ["QR", "Score"]
    
    #transpose
    df5_1 = df5.transpose()
    df5_2 = df5_1.iloc[1: , 0:]
    df5_3 = df5_2.reset_index()
    df5_4 = df5_3.iloc[0: , 1:]
    df5 = df5_4.rename(columns={2:"QR1", 3:"QR2"})
    
    # run name for x axis
    run_name = eachfile.split("/")[1].split("_")[0] + "_" + eachfile.split("/")[1].split("_")[1] 
    df5['Run'] = run_name
    
    # date 
    raw_date = str(run_name.split('_')[0])
    day = raw_date[0] + raw_date[1]
    month = raw_date[2] + raw_date[3]
    year = raw_date[4] + raw_date[5] + raw_date[6] + raw_date[7]
    df5['RawDate'] = str(year + "/" + month + "/" + day)
    df5['Date'] = pd.to_datetime(df5.RawDate)
    
    # operator and kit
    df = pd.read_excel(eachfile, sheet_name="Report")
    rows, cols = np.where(df == 'Analysed By/ Sign :')
    df2 = df.iloc[(int(rows)):(int(rows)+1), (int(cols)):(int(cols)+2)].reset_index(drop=True)
    operator = df2.iloc[0][1]
    
    rows, cols = np.where(df == 'GASTROClyr Lot Number:')
    df2 = df.iloc[(int(rows)):(int(rows)+1), (int(cols)):(int(cols)+2)].reset_index(drop=True)
    lot = df2.iloc[0][1]
    
    df5['Operator'] = operator
    df5['KitLot'] = lot
     
    # final dataframe
    final_raw = pd.concat([final_raw, df5])
    final_raw = final_raw[['Date', 'Run', 'KitLot', 'Operator', 'QR1', 'QR2']]
    
final_raw["Date"] = pd.to_datetime(final_raw["Date"])
final_raw = final_raw.sort_values(by="Date")
final_raw = final_raw[['Date', 'Run', 'KitLot', 'Operator', 'QR1', 'QR2']]
final_raw.to_csv('GCMonitoring/rawscores_qr.csv', index=False)

### Scores
scores = {}
final_scores = pd.DataFrame()

for eachfile in glob.glob("*/*.xlsm"):
    df = pd.read_excel(eachfile, sheet_name="Analysis")
    
    # find position of cell
    rows, cols = np.where(df == 'Score')
            
    # subset to only required section and clean up
    df2 = df.iloc[(int(rows[0])):(int(rows[0])+14), (int(cols[0])):(int(cols[0])+3)].reset_index(drop=True)
    # reset header and body
    df3 = df2.iloc[1:, 1:]   # start from second row and start from second column
    df3.columns = ["Sample", "Score"]
    df4 = df3[df3['Sample'] != 'NFW']   # full table
     
    # run name for x axis
    run_name = eachfile.split("/")[1].split("_")[0] + "_" + eachfile.split("/")[1].split("_")[1] 
    df4['Run'] = run_name
    
    df5 = df4.groupby('Run')['Score'].apply(lambda df4: df4.reset_index(drop=True)).unstack()
    
    # date 
    raw_date = str(run_name.split('_')[0])
    day = raw_date[0] + raw_date[1]
    month = raw_date[2] + raw_date[3]
    year = raw_date[4] + raw_date[5] + raw_date[6] + raw_date[7]
    df5['RawDate'] = str(year + "/" + month + "/" + day)
    df5['Date'] = pd.to_datetime(df5.RawDate)
    
    # operator and kit
    df = pd.read_excel(eachfile, sheet_name="Report")
    rows, cols = np.where(df == 'Analysed By/ Sign :')
    df2 = df.iloc[(int(rows)):(int(rows)+1), (int(cols)):(int(cols)+2)].reset_index(drop=True)
    operator = df2.iloc[0][1]
    
    rows, cols = np.where(df == 'GASTROClyr Lot Number:')
    df2 = df.iloc[(int(rows)):(int(rows)+1), (int(cols)):(int(cols)+2)].reset_index(drop=True)
    lot = df2.iloc[0][1]
    
    df5['Operator'] = operator
    df5['KitLot'] = lot    
    
    # final dataframe
    final_scores = pd.concat([final_scores, df5])

final_scores["Date"] = pd.to_datetime(final_scores["Date"])

# fix Run column being offset to next row
final_scores_2 = final_scores.reset_index() #reset index for Run col
final_scores_3 = final_scores_2[['Date', 'Run', 'KitLot', 'Operator', 0, 1, 2, 3, 4, 5, 6, 7, 8]]  # , 9, 10, 11, 12 taken out

final_scores_3 = final_scores_3.sort_values(by="Date")   
final_scores_3.fillna("NaN", inplace = True)
final_scores_3.to_csv('GCMonitoring/scores.csv', index=False)








  