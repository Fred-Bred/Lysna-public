# Description: This script is used to analyse the results of a team assessment survey.
# The script reads in a CSV or Excel file containing the survey data, cleans the data, calculates various scores, and generates visualizations and summary statistics for each team.
# The script also exports the results to an Excel file and a text file.

# The survey must contain the question below with specified multiple-choice options for the team selection.
# "Inden du besvarer spørgsmålene, bedes du vælge hvilket team du er en del af. Husk dette team, når du besvarer spørgsmålene. Jeg tilhører:".

# Imports
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
import re
import numpy as np
import sys
import os
import textwrap

# Load data
file_name = input("File name: ")
plots = input("Do you want to generate plots? This will add several minutes of processing (y/n): ")
assert plots in ["y", "n"], "Invalid input. Please enter 'y' or 'n'."
if plots == "y":
    dynamic_plots = input("Do you want to generate dynamic plots? This will colour plots according to normative values (y/n): ")
    assert dynamic_plots in ["y", "n"], "Invalid input. Please enter 'y' or 'n'."
    dynamic_plots = True if dynamic_plots == "y" else False

try:
    path = file_name + ".csv"
    raw_assessment = pd.read_csv(path)
except:
    try:
        path = file_name + ".xlsx"
        raw_assessment = pd.read_excel(path)
    except:
        sys.exit("File not found")

# Utility functions
# Function for sanitizing file names later
def sanitize_filename(name):
    return re.sub(r'[\/:*?"<>|]', '_', name)

# Function to wrap text
def wrap_text(text, width=65):
    return "\n".join(textwrap.wrap(text, width))

def bar_plots(data, color, output_dir, team, dynamic: bool = dynamic_plots):
    for item in data.columns.tolist():
        plt.figure(figsize=(10, 6))
        values = data[item]
        # Create a histogram with fixed bins
        counts, bins, _ = plt.hist(values, bins=np.arange(1, 6.1, 1), color=color, alpha=0)
        
        # Plot each bar individually with its own color
        if dynamic:
            for i in range(len(bins)-1):
                bin_values = values[(values >= bins[i]) & (values < bins[i+1])]
                if len(bin_values) > 0:
                    bin_color = '#ff9999' if bins[i] < 3 else '#90EE90' if bins[i] >= 4 else color
                    plt.bar(bins[i], counts[i], width=1, color=bin_color, alpha=0.5)
        else:
            plt.bar(bins[:-1], counts, width=1, color=color, alpha=0.5)
            
        plt.xticks(np.arange(1, 6, 1))
        plt.ylabel('Antal')
        wrapped_title = wrap_text(item)
        plt.title(wrapped_title)
        plt.gca().spines['top'].set_visible(False)
        plt.gca().spines['right'].set_visible(False)
        plt.gca().yaxis.set_major_locator(MaxNLocator(integer=True))
        sanitized_item = sanitize_filename(item)
        plt.savefig(f'{output_dir}/{sanitized_item}_{team}.png', transparent=True)
        plt.close()
        

def scale_plots(scale: str, team: str, dir: str, color: str = "gray", dynamic: bool = dynamic_plots) -> None:
    plt.figure(figsize=(10, 6))
    values = assessment.iloc[:-1][scale]
    # Create a histogram with fixed bins
    counts, bins, _ = plt.hist(values, bins=np.arange(1, 6.1, 1), color=color, alpha=0)
    
    # Plot each bar individually with its own color
    if dynamic:
        for i in range(len(bins)-1):
            bin_values = values[(values >= bins[i]) & (values < bins[i+1])]
            if len(bin_values) > 0:
                bin_color = '#ff9999' if bins[i] < 3 else '#90EE90' if bins[i] >= 4 else color
                plt.bar(bins[i], counts[i], width=1, color=bin_color, alpha=0.5)
    else:
        plt.bar(bins[:-1], counts, width=1, color=color, alpha=0.5)
        
    xlabel = "Organisationens Kerne" if scale == "org_core" else "Teamets Kerne" if scale == "team_core" else "Psykologisk Sikkerhed" if scale == "safety" else "Pålidelighed"
    plt.xlabel(xlabel)
    plt.xticks(np.arange(1, 6, 1))
    plt.ylabel('Antal')
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)
    plt.gca().yaxis.set_major_locator(MaxNLocator(integer=True))
    plt.savefig(f'{dir}/{scale}_{team}.png', transparent=True)
    plt.close()

def ranked_bar_plot(data: pd.DataFrame, scale: str, team: str, dir: str, color: str = "gray", dynamic: bool = dynamic_plots) -> None:
    # Calculate mean and sort
    sorted_data = data.mean().sort_values(ascending=True)
    
    # Select top 3 and bottom 3 items
    top_3 = sorted_data[-3:]
    bottom_3 = sorted_data[:3]
    
    # Concatenate top 3 and bottom 3
    combined = pd.concat([bottom_3, top_3])
    
    # Create colors array based on values
    colors = ['#ff9999' if x < 3 else '#90EE90' if x > 4 else color for x in combined]
    
    # Plot horizontal bar plot
    plt.figure(figsize=(10, 6))
    ax = plt.gca()
    if dynamic:
        bars = ax.barh(range(len(combined)), combined, color=colors)
    else:
        bars = ax.barh(range(len(combined)), combined, color=color)
    plt.xlabel('Score')
    plt.xticks(np.arange(1, 6, 1))
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)

    # Wrap y-tick labels
    y_labels = [wrap_text(label, width=40) for label in combined.index]
    ax.set_yticks(range(len(combined)))
    ax.set_yticklabels(y_labels)

    plt.tight_layout()
    plt.savefig(f'{dir}/{scale}_ranked_scores_{team}.png', transparent=True)
    plt.close()

def ranked_variance_plot(data: pd.DataFrame, scale: str, team: str, dir: str, color: str = "gray") -> None:
    # Calculate variance and sort
    sorted_data = data.var().sort_values(ascending=False)
    
    # Select top 3 and bottom 3 items
    top_3 = sorted_data.head(3)
    bottom_3 = sorted_data[-3:]
    
    # Concatenate top 3 and bottom 3
    combined = pd.concat([top_3, bottom_3])
    
    # Plot horizontal bar plot
    plt.figure(figsize=(10, 6))
    ax = combined.plot(kind='barh', color=color)
    plt.xlabel('Variance')
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)
    
    # Wrap y-tick labels
    y_labels = [wrap_text(label) for label in combined.index]
    ax.set_yticklabels(y_labels)
    
    plt.tight_layout()
    plt.savefig(f'{dir}/{scale}_variance_{team}.png', transparent=True)
    plt.close()

def get_bar_colors(values, default_color="gray", dynamic: bool = dynamic_plots):
    """Helper function to get colors based on values"""
    if not dynamic:
        return default_color
    return ['#ff9999' if x < 3 else '#90EE90' if x > 4 else default_color for x in values]


# Data cleaning
# dropping useless data
teams = ['all']
assessment = raw_assessment.drop(["Date Time", "User", "Page submitted"], axis=1)
try:
    assessment = assessment.dropna(subset="Inden du besvarer spørgsmålene, bedes du vælge hvilket team du er en del af. Husk dette team, når du besvarer spørgsmålene. Jeg tilhører:")
    mode = "multiple teams"
    teams = teams + raw_assessment["Inden du besvarer spørgsmålene, bedes du vælge hvilket team du er en del af. Husk dette team, når du besvarer spørgsmålene. Jeg tilhører:"].unique().tolist()
except:
    mode = "single team"

assessment = assessment.reset_index(drop=True)

print(f"Running analysis for {mode}.")

# Make output folders for multiple teams
if mode == "multiple teams":
    for team in teams:
        os.makedirs(team, exist_ok=True)
else:
    os.makedirs("Results", exist_ok=True)

# Run analysis for each unique team
for t in teams:
    # Data cleaning
    # dropping useless data
    team = t
    assessment = raw_assessment.drop(["Date Time", "User", "Page submitted"], axis=1)

    # Specify output dir
    if mode == "multiple teams":
        output_dir = f"{team}"
    else:
        output_dir = "Results"

    # Selecting relevant team data if not all teams
    if team != "all":
        assessment = assessment[assessment["Inden du besvarer spørgsmålene, bedes du vælge hvilket team du er en del af. Husk dette team, når du besvarer spørgsmålene. Jeg tilhører:"] == team]
    assessment = assessment.reset_index(drop=True)

    # removing " / 5"
    assessment = assessment.replace(to_replace=r"\s\/\s5", value="", regex=True)

    # Convert to numeric
    assessment.loc[:, "Jeg er bekymret for om jeg lever op til andres forventninger til mig." : "Alle i teamet føler sig personligt ansvarlige for deres opgaver."] = \
        assessment.loc[:, "Jeg er bekymret for om jeg lever op til andres forventninger til mig." : "Alle i teamet føler sig personligt ansvarlige for deres opgaver."].apply(pd.to_numeric)

    # reversing necessary items
    reversed_items = ["Jeg bliver sjældent overrasket over hvordan folk reagerer på mig og mine handlinger.",
                      "Jeg føler, at jeg kan fortælle folk mere eller mindre alt om mig.",
                      "Jeg har let ved at vise medfølelse med mine kollegaer.",
                    "Teammedlemmer ser sig selv mere som individer end som del af et team.",
                    "Hvis man begår en fejl bebrejdes man ofte for det.",
                    "Man kan nogle gange blive afvist, drillet eller ignoreret af kollegaer hvis man er anderledes.",
                    "Det er svært at bede kollegaerne i teamet om hjælp.",
                    "Teamet har tendens til at håndtere konflikter og uenigheder i mindre grupper eller udenfor teamet som helhed, i stedet for at tage dem op direkte i hele gruppen.",
                    "Drilleri i vores team er nogle gange ikke 100%kærligt ment.",
                    "I dette team gøres der ikke så god brug af mine evner som der kunne.",
                    "En / flere af kollegaerne er ikke så fagligt dygtig(e) som vi har brug for.",
                    "En / flere af kollegaerne er ikke så pålidelige som vi har brug for."]

    for item in reversed_items:
        assessment.loc[:, item] = 6 - assessment.loc[:, item]

    # Reformulate reversed items to be aligned with reversed scoring
    reformed_items = ["Jeg bliver ofte overrasket over de reaktioner jeg får fra andre.",
                      "Jeg er tilbageholdende med at fortælle andre om mig selv.",
                      "Jeg har svært ved at vise medfølelse.",
                  "Teammedlemmer ser sig selv mere som et team end som individer.",
                  "Hvis du begår en fejl på dette team, bliver det ofte tilgivet.",
                  "På dette team accepterer alle dem der er anderledes.",
                  "Det er let at bede alle i teamet om hjælp.",
                  "Teamet har tendens til at håndtere konflikter og uenigheder direkte i hele gruppen, i stedet for at tage dem op i mindre grupper eller udenfor teamet som helhed.",
                  "Drillerierne er altid 100% kærligt ment.", "Teamet gør god brug af mine evner.",
                  "Alle på teamet lever op til vores faglige standarder.", "Alle teammedlemmer er pålidelige."]

    reform_dict = dict(zip(reversed_items, reformed_items))

    assessment.rename(columns=reform_dict, inplace=True)


    # Attachment
    attachment = assessment.loc[:, "Jeg er bekymret for om jeg lever op til andres forventninger til mig." : "Jeg har svært ved at vise medfølelse."]

    # Making lists of anxiety and avoidance items
    attachment_items = attachment.columns.tolist()
    anxiety = ["Jeg er bekymret for om jeg lever op til andres forventninger til mig.",
            "Jeg bekymrer mig om hvad folk synes om mig.", "Jeg er bange for at andre ikke værdsætter mig lige så meget som jeg værdsætter dem.",
            "Jeg bliver irriteret, når jeg ikke får den støtte eller anerkendelse, som jeg føler, at jeg fortjener.",
            "Jeg bliver ofte overrasket over de reaktioner jeg får fra andre.", "Jeg må nogle gange blive vred, for at få folk til at lægge mærke til mig."]
    
    avoidance = ['Jeg forstår ikke hvorfor folk nogle gange er så følelsesladede.',
                'Jeg synes det er akavet hvis nogen på arbejdet vil betro sig til mig.',
                'Jeg er bekymret for, at hvis folk oplever hvordan jeg virkelig er, så vil de ikke kunne lide mig.',
                'Jeg foretrækker at arbejde uafhængigt af andre.',
                'Jeg vil helst ikke være for venskabelig eller tæt med mine kollegaer.',
                "Jeg er tilbageholdende med at fortælle andre om mig selv.",
                'Jeg er tilbageholden med at åbne op og vise mine følelser overfor andre.',
                "Jeg har svært ved at vise medfølelse."]

    # Add anxiety and avoidance scores
    attachment.loc[:, "anxiety"] = attachment[anxiety].sum(axis=1) / len(anxiety)
    attachment.loc[:, "avoidance"] = attachment[avoidance].sum(axis=1) / len(avoidance)
    assessment.loc[:, "anxiety"] = attachment[anxiety].sum(axis=1) / len(anxiety)
    assessment.loc[:, "avoidance"] = attachment[avoidance].sum(axis=1) / len(avoidance)

    # Round scores after old code broke and add to attachment df
    for i, score in enumerate(attachment.loc[:, "anxiety"]):
        attachment.loc[i, "anxiety"] = round(attachment.loc[i, "anxiety"], 2)

    for i, score in enumerate(attachment.loc[:, 'avoidance']):
        attachment.loc[i, 'avoidance'] = round(attachment.loc[i, 'avoidance'], 2)

    # Round scores after old code broke and add to assessment df
    for i, score in enumerate(assessment.loc[:, "anxiety"]):
        assessment.loc[i, "anxiety"] = round(assessment.loc[i, "anxiety"], 2)

    for i, score in enumerate(assessment.loc[:, 'avoidance']):
        assessment.loc[i, 'avoidance'] = round(assessment.loc[i, 'avoidance'], 2)

    # Set dtypes to float
    attachment.loc[:, 'anxiety'] = attachment['anxiety'].astype(float)
    attachment.loc[:, 'avoidance'] = attachment['avoidance'].astype(float)
    assessment.loc[:, 'anxiety'] = assessment['anxiety'].astype(float)
    assessment.loc[:, 'avoidance'] = assessment['avoidance'].astype(float)

    # Anxiety descriptive statistics
    round(assessment.loc[:"mean", "anxiety"].describe()[["mean", "min", "max"]], 2)

    # Store anxiety scores in text file
    with open(f'{output_dir}/Assessment results_{team}.txt', 'w', encoding='utf-8') as f:
        f.write(f'Team anxiety scores: \n{round(assessment.loc[:"mean", "anxiety"].describe()[["mean", "min", "max"]], 2)}')

    # Avoidance descriptive statistics
    round(assessment.loc[:"mean", "avoidance"].describe()[["mean", "min", "max"]], 2)

    # Store avoidance scores in text file
    with open(f'{output_dir}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
        f.write(f'\n \nTeam avoidance scores: \n{round(assessment.loc[:"mean", "avoidance"].describe()[["mean", "min", "max"]], 2)}')

    # Plot attachment scores
    fig, ax = plt.subplots(figsize=(10, 7), dpi=400)

    ### Old method with same marker for every point
    # Colouring and plotting with different colours but same markers (uncomment to use)
    # colors = cm.rainbow(np.linspace(0, 1, len(attachment["avoidance"])))
    # ax.scatter(attachment["anxiety"], attachment["avoidance"], color=colors, alpha=0.5) # plotting avoidance against attachment with semi-transparent dots and unique colours to display similar values stacked


    ### New method with different marker for each point
    ## (probably) Redundant code to set figsize and layout for all figures:
    # plt.rcParams["figure.figsize"] = [7.50, 3.50]
    # plt.rcParams["figure.autolayout"] = True

    x = attachment['avoidance']
    y = attachment['anxiety']
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
    # plt.text(4.42, 4.10, "Frygt/dominans", fontsize=18, color="darkgray", alpha=0.35)

    fig.set_facecolor("white") # set background colour to white to make labels and title visible in dark mode

    # save and show
    plt.savefig(f"{output_dir}/Attachment_plot_{team}.png", transparent=True)
    plt.close()

    # Add mean to each item and overall score in attachment dataframe
    attachment.loc["mean"] = attachment.mean()

    # Rounding
    for col in list(attachment.columns):
        attachment.loc["mean", col] = round(attachment.loc['mean', col], 2)

    # Organisational core
    org_core = assessment.loc[:, "Vores arbejdsgiver/virksomhed har et tydeligt værdisæt (f.eks. omkring godt samarbejde og hvordan vi opfører os overfor hinanden) som alle følger på denne plads/ i dette projekt.":"Alle involverede faggrupper (også udenfor vores virksomhed) arbejder gnidningsfrit sammen om at færdiggøre projektet."]

    # Calculating "core" score and adding to core and overall DataFrames
    # creating list of variables
    org_core_items = org_core.columns.tolist()

    # adding total core score to org_core DataFrame
    org_core.loc[:, "org_core"] = org_core.loc[:, org_core_items].sum(axis=1) / len(org_core_items)

    # Round all core scores
    org_core.loc[:, "org_core"] = org_core["org_core"].round(2)

    # Set dtype of core column
    org_core.loc[:, "org_core"] = org_core["org_core"].astype(float)

    # Add mean scores
    org_core.loc["mean"] = org_core[org_core_items].mean()

    # Add rounded mean to each column
    for col in list(org_core.columns):
        org_core.loc["mean", col] = round(org_core.loc['mean', col], 2)

    # Add core mean
    org_core.loc['mean', "org_core"] = org_core.loc[:"mean", "org_core"].mean()

    # Add core scores to assessment df
    assessment.loc[:, "org_core"] = org_core.loc[:, "org_core"]

    # Create numeric df
    org_core_numeric = org_core.iloc[:-1, :-1].select_dtypes(include='number')

    # Variance
    org_core_sorted_var = org_core_numeric.var().sort_values(ascending=False).round(2)
    org_core_top_var = org_core_sorted_var[:3]
    org_core_bottom_var = org_core_sorted_var[-3:]

    # # Variance for each item in core
    # org_core_variance = {}
    # for col in org_core_items:
    #     org_core_variance[col] = round(org_core.drop('mean', axis=0)[col].var(), 2)

    # org_core_variance = {k: v for k, v in sorted(org_core_variance.items(), key=lambda item: item[1])}

    # # 3 lowest variance items
    # org_core_low_list = list(org_core_variance.keys())[:3]
    # org_core_low_var = org_core.loc['mean', org_core_low_list]

    # # 3 highest variance items
    # org_core_high_list = list(org_core_variance.keys())[-3:]

    # org_core_high_var = org_core.loc['mean', org_core_high_list]

    # Write organisational core stats to text file
    with open(f'{output_dir}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
        # Descriptive stats
        f.write(f'\n \nOrganisational core: \n{round(org_core.loc[:"mean", "org_core"].describe()[["mean", "min", "max"]], 2)}')
        # Lowest and highest scoring items
        f.write(f'\n \nThree lowest scoring org_core items: \n{org_core.iloc[-1, :-1].sort_values()[:3]}\n \nThree highest scoring org_core items: \n{org_core.iloc[-1, :-1].sort_values(ascending=False)[:3]}')
        ## Lowest and highest variance items
        f.write(f'\n \nThree lowest variance org_core items: \n{org_core_bottom_var} \n \nThree highest variance org_core items: \n{org_core_top_var}')

    # Team safety
    safety = assessment.loc[:, "Teammedlemmer ser sig selv mere som et team end som individer.":"Drillerierne er altid 100% kærligt ment."]

    # Calculating safety score and adding to overall df
    # creating list of variables
    safety_items = safety.columns.tolist()

    # adding total safety score to safety DataFrame
    safety.loc[:, "safety"] = safety.loc[:, safety_items].sum(axis=1) / len(safety_items)

    # Round all safety score
    safety.loc[:, "safety"] = safety["safety"].round(2)

    # Set dtype of safety column
    safety.loc[:, 'safety'] = safety['safety'].astype(float)

    # Add mean scores
    safety.loc["mean"] = safety[safety_items].mean()

    # Add safety mean
    safety.loc['mean', 'safety'] = safety.loc[:"mean", "safety"].mean()

    # Add rounded mean to each column
    for col in list(safety.columns):
        safety.loc["mean", col] = round(safety.loc['mean', col], 2)

    # adding total core score to assessment dataframe
    assessment.loc[:, "safety"] = safety.loc[:, 'safety']

    # Create numeric df
    safety_numeric = safety.iloc[:-1, :-1].select_dtypes(include='number')

    # Variance
    safety_sorted_var = safety_numeric.var().sort_values(ascending=False).round(2)
    safety_top_var = safety_sorted_var[:3]
    safety_bottom_var = safety_sorted_var[-3:]

    # # Variance for each safety item
    # safety_variance = {}
    # for col in safety_items:
    #     safety_variance[col] = round(safety.drop('mean', axis=0)[col].var(), 2)

    # safety_variance = {k: v for k, v in sorted(safety_variance.items(), key=lambda item: item[1])}

    # # 3 lowest variance items
    # safety_low_list = list(safety_variance.keys())[:3]
    # safety_low_var = safety.loc['mean', safety_low_list]

    # # 3 highest variance items
    # safety_high_list = list(safety_variance.keys())[-3:]
    # safety_high_var = safety.loc['mean', safety_high_list]

    # Write safety stats to text file
    with open(f'{output_dir}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
        # Descriptive stats
        f.write(f'\n \nTeam safety: \n{round(safety.loc[:"mean", "safety"].describe()[["mean", "min", "max"]], 2)}')
        # Lowest and highest scoring items
        f.write(f'\n \nThree lowest scoring safety items: \n{safety.iloc[-1, :-1].sort_values()[:3]}\n \nThree highest scoring safety items: \n{safety.iloc[-1, :-1].sort_values(ascending=False)[:3]}')
        ## Lowest and highest variance items
        f.write(f'\n \nThree lowest variance safety items: \n{safety_bottom_var} \n \nThree highest variance safety items: \n{safety_top_var}')

    # Team dependability
    dependability = assessment.loc[:, "Teamet har klare fælles mål.":"Alle i teamet føler sig personligt ansvarlige for deres opgaver."]

    # Calculating dependability score and adding to overall df
    # list of items
    dependability_items = dependability.columns.tolist()

    # adding total dependability score to dependability DataFrame
    dependability.loc[:, "dependability"] = dependability.loc[:, dependability_items].sum(axis=1) / len(dependability_items)

    # Round all dependability score
    dependability.loc[:, "dependability"] = dependability["dependability"].round(2)

    # Set dtype of dependability column
    dependability.loc[:, 'dependability'] = dependability['dependability'].astype(float)

    # Add mean scores
    dependability.loc["mean"] = dependability[dependability_items].mean()

    # Add dependability mean
    dependability.loc['mean', 'dependability'] = dependability.loc[:"mean", "dependability"].mean()

    # Add rounded mean to each column
    for col in list(dependability.columns):
        dependability.loc["mean", col] = round(dependability.loc['mean', col], 2)

    # adding total dependability score to assessment dataframe
    assessment.loc[:, "dependability"] = dependability.loc[:, 'dependability']

    # Create numeric df
    dependability_numeric = dependability.iloc[:-1, :-1].select_dtypes(include='number')

    # Variance
    dependability_sorted_var = dependability_numeric.var().sort_values(ascending=False).round(2)
    dependability_top_var = dependability_sorted_var[:3]
    dependability_bottom_var = dependability_sorted_var[-3:]

    # # Variance for each item in dependability
    # dependability_variance = {}
    # for col in dependability_items:
    #     dependability_variance[col] = round(dependability.drop('mean', axis=0)[col].var(), 2)

    # dependability_variance = {k: v for k, v in sorted(dependability_variance.items(), key=lambda item: item[1])}

    # # 3 lowest variance items
    # dependability_low_list = list(dependability_variance.keys())[:3]
    # dependability_low_var = dependability.loc['mean', dependability_low_list]

    # # 3 highest variance items
    # dependability_high_list = list(dependability_variance.keys())[-3:]
    # dependability_high_var = dependability.loc['mean', dependability_high_list]

    # Write dependability stats to text file
    with open(f'{output_dir}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
        # Descriptive stats
        f.write(f'\n \nDependability: \n{round(dependability.loc[:"mean", "dependability"].describe()[["mean", "min", "max"]], 2)}')
        # Lowest and highest scoring items
        f.write(f'\n \nThree lowest scoring dependability items: \n{dependability.iloc[-1, :-1].sort_values()[:3]}\n \nThree highest scoring dependability items: \n{dependability.iloc[-1, :-1].sort_values(ascending=False)[:3]}')
        ## Lowest and highest variance items
        f.write(f'\n \nThree lowest variance dependability items: \n{dependability_bottom_var} \n \nThree highest variance dependability items: \n{dependability_top_var}')

    # Final DataFrame
    # Add rounded mean to each column
    for col in list(assessment.columns):
        try:
            assessment.loc["mean", col] = round(assessment[col].mean(), 2)
        except:
            continue

    # Exporting to Excel
    assessment.to_excel(f"{output_dir}/Team assessment results {team}.xlsx")

    ### PLOTTING (optional)
    if plots == "n":
        continue # End this loop iteration if plots are not requested

    print(f"Plotting for {team}...")

    # # org_core
    org_core_dir = f"{output_dir}/org_core"
    os.makedirs(org_core_dir, exist_ok=True)

    scale_plots("org_core", team, org_core_dir, color="gray")

    # ## Highest scoring org_core items
    # bar_plots(org_core.iloc[-1, :-1].sort_values(ascending=False)[:3].index.tolist(),
    #                org_core, "gray", org_core_dir, team)

    # ## Lowest scoring org_core items
    # bar_plots(org_core.iloc[-1, :-1].sort_values()[:3].index.tolist(),
    #                org_core, "gray", org_core_dir, team)

    # # Highest variance org_core items
    # bar_plots(org_core_high_list, org_core, "gray", org_core_dir, team)

    # # Lowest variance org_core items
    # bar_plots(org_core_low_list, org_core, "gray", org_core_dir, team)

    # Plot all numeric org_core items
    bar_plots(org_core_numeric, "gray", org_core_dir, team)
    ranked_bar_plot(org_core_numeric, "org_core", team, org_core_dir)
    ranked_variance_plot(org_core_numeric, "org_core", team, org_core_dir)

    # safety
    safety_dir = f"{output_dir}/safety"
    os.makedirs(safety_dir, exist_ok=True)

    scale_plots("safety", team, safety_dir, color="gray")

    # ## Highest scoring safety items
    # bar_plots(safety.iloc[-1, :-1].sort_values(ascending=False)[:3].index.tolist(),
    #                     safety, "gray", safety_dir, team)

    # ## Lowest scoring safety items
    # bar_plots(safety.iloc[-1, :-1].sort_values()[:3].index.tolist(),
    #                     safety, "gray", safety_dir, team)

    # # Highest variance safety items
    # bar_plots(safety_high_list, safety, "gray", safety_dir, team)

    # # Lowest variance safety items
    # bar_plots(safety_low_list, safety, "gray", safety_dir, team)

    # Plot all numeric safety items
    bar_plots(safety_numeric, "gray", safety_dir, team)
    ranked_bar_plot(safety_numeric, "safety", team, safety_dir)
    ranked_variance_plot(safety_numeric, "safety", team, safety_dir)

    # # dependability
    dependability_dir = f"{output_dir}/dependability"
    os.makedirs(dependability_dir, exist_ok=True)

    scale_plots("dependability", team, dependability_dir, color="gray")

    # # Highest scoring dependability items
    # bar_plots(dependability.iloc[-1, :-1].sort_values(ascending=False)[:3].index.tolist(),
    #                     dependability, "gray", dependability_dir, team)

    # ## Lowest scoring dependability items
    # bar_plots(dependability.iloc[-1, :-1].sort_values()[:3].index.tolist(),
    #                     dependability, "gray", dependability_dir, team)

    # # Highest variance dependability items
    # bar_plots(dependability_high_list, dependability, "gray", dependability_dir, team)

    # # Lowest variance dependability items
    # bar_plots(dependability_low_list, dependability, "gray", dependability_dir, team)

    # Plot all numeric dependability items
    bar_plots(dependability_numeric, "gray", dependability_dir, team)
    ranked_bar_plot(dependability_numeric, "dependability", team, dependability_dir)
    ranked_variance_plot(dependability_numeric, "dependability", team, dependability_dir)

    # Plot variance for all scales
    scales = ["org_core", "safety", "dependability"]
    plt.figure(figsize=(10, 6))
    plt.bar(assessment[scales].var().index, assessment[scales].var(), color="gray")
    plt.xticks(ticks=range(len(scales)), labels=["Organisationens Kerne", "Psykologisk Sikkerhed", "Pålidelighed"])
    plt.tight_layout()
    plt.savefig(f'{output_dir}/variance_{team}.png', transparent=True)
    plt.close()

    # Plot scores for all scales with conditional coloring
    plt.figure(figsize=(10, 6))
    values = assessment.loc["mean", scales]
    colors = get_bar_colors(values) if dynamic_plots else "gray"
    plt.bar(range(len(scales)), values, color=colors)
    plt.yticks(ticks=np.arange(1, 6, 1))
    plt.xticks(ticks=range(len(scales)), labels=["Organisationens Kerne", "Psykologisk Sikkerhed", "Pålidelighed"])
    plt.tight_layout()
    plt.savefig(f'{output_dir}/scale_scores_{team}.png', transparent=True)
    plt.close()

print("Analysis complete. Results exported to Excel and text files.")