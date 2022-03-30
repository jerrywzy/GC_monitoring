import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import glob

## dCt of PCR replicates 
dct = {}

for gene_count in range(1,13):
    for eachfile in glob.glob("*/*.xlsm"):
        df = pd.read_excel(eachfile, sheet_name="Analysis")
        
        # find position of cell
        rows, cols = np.where(df == 'dCt of PCR replicates')
        
        # subset to only required section and clean up
        df2 = df.iloc[(int(rows)):(int(rows)+16), (int(cols)):(int(cols)+13)].reset_index(drop=True)
        # reset header and body
        df2.columns = df2.iloc[0]
        df3 = df2[1:]
        
        # gene number for graph titles
        gene_number = df3.columns[gene_count]     
        # run name for x axis
        run_name = eachfile.split("\\")[1].split("_")[0] + "_" + eachfile.split("\\")[1].split("_")[1]
            
        # add to dictionary to create dataframe
        y = df3[df3.columns[gene_count]]
        dct[run_name] = y
    
    # create dataframe and plot
    dct_df = pd.DataFrame(dct)
    
    # create a figure
    fig, ax = plt.subplots()
    ax = sns.stripplot(data=dct_df)
    ax.set(title=gene_number + " PCR Replicates")
    ax.set_xticklabels(ax.get_xticklabels(), rotation=90, ha="right")
    plt.tight_layout()
    plt.savefig("GCMonitoring/PCRReplicates_{0}.png".format(gene_number))

### NC
nc = {}

for gene_count in range(1,13):
    for eachfile in glob.glob("*/*.xlsm"):
        df = pd.read_excel(eachfile, sheet_name="Analysis")
        
        # find position of cell
        rows, cols = np.where(df == 'Average of PCR replicates')
        
        # subset to only required section and clean up
        df2 = df.iloc[(int(rows)):(int(rows)+17), (int(cols)):(int(cols)+14)].reset_index(drop=True)
        # reset header and body
        df2.columns = df2.iloc[0]
        df2_2 = df2.rename(columns={np.nan: 'Average of PCR replicates'})
        df3 = df2_2[1:]
        df3_2 = df3.iloc[[-1], 1:]   # take last row and start from second column
        df3_3 = df3_2.fillna(40)
        
        # gene number for graph titles
        gene_number = df3_3.columns[gene_count] 
        # run name for x axis
        run_name = eachfile.split("\\")[1].split("_")[0] + "_" + eachfile.split("\\")[1].split("_")[1]
        
        # add to dictionary to create dataframe
        y = df3_3[df3_3.columns[gene_count]]
        nc[run_name] = y
    
    # create dataframe and plot
    nc_df = pd.DataFrame(nc)
    
    # create a figure
    fig, ax = plt.subplots()   
    ax = sns.stripplot(data=nc_df)
    ax.set(title=gene_number + " NC")
    ax.set_xticklabels(ax.get_xticklabels(), rotation=90, ha="right")
    plt.tight_layout()
    plt.savefig("GCMonitoring/NC_{0}.png".format(gene_number))


### Raw Score
raw_score = {}

for eachfile in glob.glob("*/*.xlsm"):
    df = pd.read_excel(eachfile, sheet_name="Analysis")
    
    # find position of cell
    rows, cols = np.where(df == 'Raw score')
    
    # subset to only required section and clean up
    df2 = df.iloc[(int(rows)):(int(rows)+4), (int(cols)):(int(cols)+3)].reset_index(drop=True)
    # reset header and body
    df3 = df2.iloc[[2,3], 1:]   # take second and third row and start from second column
    
    # run name for x axis
    run_name = eachfile.split("\\")[1].split("_")[0] + "_" + eachfile.split("\\")[1].split("_")[1] 
     
    # add to dictionary to create dataframe
    y = df3[df3.columns[1]]
    raw_score[run_name] = y

# create dataframe and plot
raw_score_df = pd.DataFrame(raw_score)

# create a figure
fig, ax = plt.subplots()
ax = sns.stripplot(data=raw_score_df)
ax.set(title="QR Scores")
ax.set_xticklabels(ax.get_xticklabels(), rotation=90, ha="right")
plt.tight_layout()

# create mean lines
sns.boxplot(showmeans=True,
            meanline=True,
            meanprops={'color': 'k', 'ls': '-', 'lw': 2},
            medianprops={'visible': False},
            whiskerprops={'visible': False},
            zorder=10,
            data=raw_score_df,
            showfliers=False,
            showbox=False,
            showcaps=False,
            ax=ax)

plt.savefig("GCMonitoring/QRScores.png")


### Scores
scores = {}

for eachfile in glob.glob("*/*.xlsm"):
    df = pd.read_excel(eachfile, sheet_name="Analysis")
    
    # find position of cell
    rows, cols = np.where(df == 'Score')
            
    # subset to only required section and clean up
    df2 = df.iloc[(int(rows[0])):(int(rows[0])+14), (int(cols[0])):(int(cols[0])+3)].reset_index(drop=True)
    # reset header and body
    df3 = df2.iloc[1:, 1:]   # start from second row and start from second column
     
    # run name for x axis
    run_name = eachfile.split("\\")[1].split("_")[0] + "_" + eachfile.split("\\")[1].split("_")[1] 
         
    # add to dictionary to create dataframe
    y = df3[df3.columns[1]]
    scores[run_name] = y

# create dataframe and plot
scores_df = pd.DataFrame(scores)

# create a figure
fig, ax = plt.subplots()  
ax = sns.stripplot(data=scores_df)
ax.set(title="Average Scores across all samples")
ax.set_xticklabels(ax.get_xticklabels(), rotation=90, ha="right")
plt.tight_layout()

# create mean lines
sns.boxplot(showmeans=True,
            meanline=True,
            meanprops={'color': 'k', 'ls': '-', 'lw': 2},
            medianprops={'visible': False},
            whiskerprops={'visible': False},
            zorder=10,
            data=scores_df,
            showfliers=False,
            showbox=False,
            showcaps=False,
            ax=ax)

plt.savefig("GCMonitoring/AverageScores.png")



