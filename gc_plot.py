import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import glob

### dct
dct_df = pd.read_csv('GCMonitoring/dct.csv')
    
for gene_number in range(1,13):
    
    col_name = ("G" + str(gene_number))
    
    # subset to G1-G12 each
    df_to_plot_pre = dct_df[['Run', 'dCt of PCR replicates', col_name]]
    df_to_plot = df_to_plot_pre.drop(df_to_plot_pre[df_to_plot_pre['dCt of PCR replicates'] == "NC"].index)  # to drop NC, make sure no samples are >30 also
    
    # create a figure
    fig, ax = plt.subplots(figsize=(12.7, 8.27))
    ax = sns.stripplot(x='Run', y=col_name, data=df_to_plot)
    ax.set(title=str(col_name) + " PCR Replicates")
    ax.set_xticklabels(ax.get_xticklabels(), rotation=90, ha="right")
    plt.tight_layout()
    plt.savefig("GCMonitoring/PCRReplicates_{0}.png".format(col_name))

### NC
nc_df = pd.read_csv('GCMonitoring/nc.csv')
    
for gene_number in range(1,13):
    
    col_name = ("G" + str(gene_number))
    
    # subset to G1-G12 each
    df_to_plot_pre = nc_df[['Run', 'Average of PCR replicates', col_name]]
    df_to_plot = df_to_plot_pre.drop(df_to_plot_pre[df_to_plot_pre['Average of PCR replicates'] != "NC"].index)  # to drop NC, make sure no samples are >30 also
    df_to_plot.to_csv('GCMonitoring/nc_only.csv', index=False)
    
    # create a figure
    fig, ax = plt.subplots(figsize=(12.7, 8.27))
    ax = sns.stripplot(x='Run', y=col_name, data=df_to_plot)
    ax.set(title=str(col_name) + " NC")
    ax.set_xticklabels(ax.get_xticklabels(), rotation=90, ha="right")
    plt.tight_layout()
    plt.savefig("GCMonitoring/NC_{0}.png".format(col_name))
    
### Raw Score (WIP Stacking)
raw_df = pd.read_csv('GCMonitoring/rawscores_qr.csv')

# subset to specific columns
df_to_plot = raw_df[['Run', 'QR1', 'QR2']]
dfm = df_to_plot.melt('Run', var_name='QR', value_name='Value')

# create a figure
fig, ax = plt.subplots(figsize=(12.7, 8.27))
ax = sns.stripplot(x='Run', y='Value', data=dfm)
ax.set(title="QR scores")
ax.set_xticklabels(ax.get_xticklabels(), rotation=90, ha="right")
plt.tight_layout()

# plot the mean line
sns.boxplot(showmeans=True,
            meanline=True,
            meanprops={'color': 'k', 'ls': '-', 'lw': 2},
            medianprops={'visible': False},
            whiskerprops={'visible': False},
            zorder=10,
            x="Run",
            y="Value",
            data=dfm,
            showfliers=False,
            showbox=False,
            showcaps=False,
            ax=ax)

plt.savefig("GCMonitoring/QRScores.png")

### Scores
scores_df = pd.read_csv('GCMonitoring/scores.csv')
    
# subset to scores up to sample 9 only
df_to_plot = scores_df[['Run', '0', '1', '2', '3', '4', '5', '6', '7', '8']]
# df_to_plot['Average'] = df_to_plot.iloc[:, 1:9].mean(axis=1)
dfm = df_to_plot.melt('Run', var_name='Sample', value_name='Value')

# create a figure
fig, ax = plt.subplots(figsize=(12.7, 8.27))
ax = sns.stripplot(x='Run', y='Value', data=dfm)
ax.set(title="Average Scores across all samples")
ax.set_xticklabels(ax.get_xticklabels(), rotation=90, ha="right")
plt.tight_layout()

# plot the mean line
sns.boxplot(showmeans=True,
            meanline=True,
            meanprops={'color': 'k', 'ls': '-', 'lw': 2},
            medianprops={'visible': False},
            whiskerprops={'visible': False},
            zorder=10,
            x="Run",
            y="Value",
            data=dfm,
            showfliers=False,
            showbox=False,
            showcaps=False,
            ax=ax)

plt.savefig("GCMonitoring/AverageScores.png")


