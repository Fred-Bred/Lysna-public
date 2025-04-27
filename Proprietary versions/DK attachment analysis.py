# Imports
import pandas as pd
import matplotlib.pyplot as plt
import re
import numpy as np
import sys
import os
import textwrap

# Load data
file_name = input("File name: ")

try:
    path = file_name + ".csv"
    raw_assessment = pd.read_csv(path)
except:
    try:
        path = file_name + ".xlsx"
        raw_assessment = pd.read_excel(path)
    except:
        sys.exit("File not found")

# Data cleaning
# dropping useless data
teams = ['all']
assessment = raw_assessment.drop(["Date Time", "User", "Page submitted"], axis=1)

# removing " / 5"
assessment = assessment.replace(to_replace=r"\s\/\s5", value="", regex=True)

# Storing values as integers
assessment.loc[:, "Jeg er bekymret for om jeg lever op til andres forventninger til mig." : "Jeg har let ved at vise medfølelse."] = \
    assessment.loc[:, "Jeg er bekymret for om jeg lever op til andres forventninger til mig." : "Jeg har let ved at vise medfølelse."].apply(pd.to_numeric)

# reversing necessary items
reversed_items = ["Jeg bliver sjældent overrasket over hvordan folk reagerer på min fremtræden eller adfærd.", "Jeg føler, at jeg kan fortælle folk mere eller mindre alt om mig.", "Jeg har let ved at vise medfølelse."]

for item in reversed_items:
    assessment.loc[:, item] = 6 - assessment.loc[:, item]

# Reformulate reversed items to be aligned with reversed scoring
reformed_items = ["Jeg bliver ofte overrasket over de reaktioner jeg får fra andre.", "Jeg er tilbageholdende med at fortælle andre om mig selv.", "Jeg har svært ved at vise medfølelse."]

reform_dict = dict(zip(reversed_items, reformed_items))

assessment.rename(columns=reform_dict, inplace=True)

# Make lists of items for each factor
anxiety = ["Jeg er bekymret for om jeg lever op til andres forventninger til mig.",
        "Jeg bekymrer mig om hvad folk synes om mig.", "Jeg er bange for at andre ikke værdsætter mig så meget som jeg værdsætter dem.",
        "Jeg bliver irriteret, når jeg ikke får den støtte eller anerkendelse, som jeg føler, at jeg fortjener.",
        "Jeg bliver ofte overrasket over de reaktioner jeg får fra andre.", "Jeg må nogle gange blive vred, for at få folk til at lægge mærke til mig."]
avoidance = ['Jeg forstår ikke hvorfor folk nogle gange er så følelsesladede.',
            'Jeg bliver ukomfortabel hvis nogen på arbejdet vil betro sig til mig.',
            'Jeg er bange for, at hvis folk ser hvordan jeg virkelig er, så vil de reagere negativt eller ikke kunne lide mig.',
            'Jeg foretrækker at arbejde uafhængigt af andre.',
            'Jeg vil helst ikke være for venskabelig eller tæt med mine kollegaer.',
            "Jeg er tilbageholdende med at fortælle andre om mig selv.",
            'Jeg er tilbageholden med at åbne op overfor andre.',
            "Jeg har svært ved at vise medfølelse."]

# Add anxiety and avoidance scores
assessment.loc[:, "anxiety"] = assessment[anxiety].sum(axis=1) / len(anxiety)
assessment.loc[:, "avoidance"] = assessment[avoidance].sum(axis=1) / len(avoidance)

# Round scores after old code broke
for i, score in enumerate(assessment.loc[:, "anxiety"]):
    assessment.loc[i, "anxiety"] = round(assessment.loc[i, "anxiety"], 2)

for i, score in enumerate(assessment.loc[:, 'avoidance']):
    assessment.loc[i, 'avoidance'] = round(assessment.loc[i, 'avoidance'], 2)

# Set dtypes to float'
assessment.loc[:, 'anxiety'] = assessment['anxiety'].astype(float)
assessment.loc[:, 'avoidance'] = assessment['avoidance'].astype(float)

# Plot attachment scores for all members in one plot
fig, ax = plt.subplots(figsize=(10, 7), dpi=400)

### Old method with same marker for every point
# Colouring and plotting with different colours but same markers (uncomment to use)
# colors = cm.rainbow(np.linspace(0, 1, len(attachment["avoidance"])))
# ax.scatter(attachment["anxiety"], attachment["avoidance"], color=colors, alpha=0.5) # plotting avoidance against attachment with semi-transparent dots and unique colours to display similar values stacked


### New method with different marker for each point
## (probably) Redundant code to set figsize and layout for all figures:
# plt.rcParams["figure.figsize"] = [7.50, 3.50]
# plt.rcParams["figure.autolayout"] = True

x = assessment['avoidance']
y = assessment["anxiety"]
markers = ["d", "v", "s", "*", "^", "*", "^", "x", "p", "P", "D", "H", "1", "2", "3", "4", "<", ">"]
for xp, yp, m in zip(x, y, markers):
    plt.scatter(xp, yp, marker=m, s=50, alpha=0.75)

### Shared code to make the plot pretty and self-contained
# Reverse axes
plt.gca().invert_xaxis()
plt.gca().invert_yaxis()
# labels
# plt.title("Culture")
plt.xlabel("Tillid")
plt.ylabel("Mod")
# x and y ticks
plt.xticks([1, 2, 3, 4, 5])
plt.yticks([1, 2, 3, 4, 5])

plt.tick_params(
    axis='x',          # changes apply to the x-axis
    which='major',      # both major and minor ticks are affected
    bottom=False,      # ticks along the bottom edge are off
    top=False,         # ticks along the top edge are off
    labelbottom=False) # labels along the bottom edge are off

plt.tick_params(
    axis='y',          # changes apply to the x-axis
    which='major',      # both major and minor ticks are affected
    left=False,      # ticks along the bottom edge are off
    right=False,         # ticks along the top edge are off
    labelleft=False) # labels along the bottom edge are off

# lines
ax.axhline(3, color="gray", linewidth=1, ls="--", alpha=0.6) # add horizontal line
ax.axvline(3, color="gray", linewidth=1, ls="--", alpha=0.6) # add vertical line
# text
plt.text(2.45, 2.10, "Transformation", fontsize=18, color="darkgray")
plt.text(4.25, 2.10, "Dominans", fontsize=18, color="darkgray")
plt.text(2.25, 4.10, "Omsorg", fontsize=18, color="darkgray")
# plt.text(4.42, 4.10, "Frygt/dominans", fontsize=18, color="darkgray")

fig.set_facecolor("white") # set background colour to white to make labels and title visible in dark mode

# save and show
plt.savefig(f"Attachment_plot.png")
plt.close()

# Produce unique plots for each individual
for i, name in enumerate(assessment["Mit (kalde) navn er:"]):
    fig, ax = plt.subplots(figsize=(10, 7), dpi=400)
    plt.scatter(assessment.loc[i, "avoidance"], assessment.loc[i, "anxiety"], marker="o", s=50, alpha=0.75)

    # Invert axes
    plt.gca().invert_xaxis()
    plt.gca().invert_yaxis()

    # Axis labels
    plt.xlabel("Tillid")
    plt.ylabel("Mod")
    plt.xticks([1, 2, 3, 4, 5])
    plt.yticks([1, 2, 3, 4, 5])
    
    # Remove ticks
    plt.tick_params(
        axis='x',
        which='major',
        bottom=False,
        top=False,
        labelbottom=False)
    
    plt.tick_params(
        axis='y',
        which='major',
        left=False,
        right=False,
        labelleft=False)
    
    # Add lines and text
    ax.axhline(3, color="gray", linewidth=1, ls="--", alpha=0.6)
    ax.axvline(3, color="gray", linewidth=1, ls="--", alpha=0.6)
    
    plt.text(2.45, 2.10, "Transformation", fontsize=18, color="darkgray")
    plt.text(4.25, 2.10, "Dominans", fontsize=18, color="darkgray")
    plt.text(2.25, 4.10, "Omsorg", fontsize=18, color="darkgray")
    # plt.text(4.42, 4.10, "Frygt/dominans", fontsize=18, color="darkgray")
    
    # Save and close
    fig.set_facecolor("white")
    plt.savefig(f"Attachment_plot_{name}.png")
    plt.close()

# Add mean to each item and overall score
assessment.loc["mean"] = assessment.mean(numeric_only=True)

# Save data
assessment.to_excel("attachment_scores.xlsx")
