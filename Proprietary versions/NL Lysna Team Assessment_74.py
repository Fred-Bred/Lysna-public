# Description: This script is used to analyse the results of a team assessment survey.
# The script reads in a CSV or Excel file containing the survey data, cleans the data, calculates various scores, and generates visualizations and summary statistics for each team.
# The script also exports the results to an Excel file and a text file.

# The survey must contain the question below with specified multiple-choice options for the team selection.
# "Voor het beantwoorden van de vragen, selecteer eerst bij welk team je hoort. Houd dit team in gedachten bij het beantwoorden van de vragen. Ik hoor bij team:".

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

try:
    path = file_name + ".csv"
    raw_assessment = pd.read_csv(path)
except:
    try:
        path = file_name + ".xlsx"
        raw_assessment = pd.read_excel(path)
    except:
        sys.exit("File not found, Please try again.")

# Utility functions
# Function for sanitizing file names later
def sanitize_filename(name):
    return re.sub(r'[\/:*?"<>|]', '_', name)

# Function to wrap text
def wrap_text(text, width=65):
    return "\n".join(textwrap.wrap(text, width))

def bar_plots(data, color, output_dir, team):
    for item in data.columns.to_list():
        plt.figure(figsize=(10, 6))
        plt.hist(data[item], bins=10, color=color, alpha=0.5)
        # plt.xlabel(item)
        plt.xticks(np.arange(1, 6, 1))
        plt.ylabel('Aantal')
        wrapped_title = wrap_text(item)
        plt.title(wrapped_title)
        plt.gca().spines['top'].set_visible(False)
        plt.gca().spines['right'].set_visible(False)
        # Ensure integer y-axis
        plt.gca().yaxis.set_major_locator(MaxNLocator(integer=True))
        sanitized_item = sanitize_filename(item)
        plt.savefig(f'{output_dir}/{sanitized_item}_{team}.png', transparent=True)
        plt.close()


def scale_plots(scale: str, team: str, dir: str, color: str = "gray") -> None:
    plt.figure(figsize=(10, 6))
    plt.hist(assessment.iloc[:-1][scale], bins=10, color=color, alpha=0.5)
    xlabel = "Kern van de organisatie" if scale == "org_core" else "Team kern" if scale == "team_core" else "Team veiligheid" if scale == "safety" else "Team betrouwbaarheid"
    plt.xlabel(xlabel)
    plt.xticks(np.arange(1, 6, 1))
    plt.ylabel('Aantal')
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)
    plt.gca().yaxis.set_major_locator(MaxNLocator(integer=True))
    plt.savefig(f'{dir}/{scale}_{team}.png', transparent=True)
    plt.close()

def ranked_bar_plot(data: pd.DataFrame, scale: str, team: str, dir: str, color: str = "gray") -> None:
    # Calculate mean and sort
    sorted_data = data.mean().sort_values(ascending=True)
    
    # Select top 3 and bottom 3 items
    top_3 = sorted_data[-3:]
    bottom_3 = sorted_data[:3]
    
    # Concatenate top 3 and bottom 3
    combined = pd.concat([bottom_3, top_3])
    
    # Plot horizontal bar plot
    plt.figure(figsize=(10, 6))
    ax = combined.plot(kind='barh', color=color)
    plt.xlabel('Score')
    plt.xticks(np.arange(1, 6, 1))
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)

    # Wrap y-tick labels
    y_labels = [wrap_text(label, width=40) for label in combined.index]
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


# Data cleaning
# dropping useless data
assessment = raw_assessment.drop(["Date Time", "User", "Page submitted"], axis=1)
teams = ["all"]
try:
    assessment = assessment.dropna(subset="Voor het beantwoorden van de vragen, selecteer eerst bij welk team je hoort. Houd dit team in gedachten bij het beantwoorden van de vragen. Ik hoor bij team:")
    mode = "multiple teams"
    teams = teams + raw_assessment["Voor het beantwoorden van de vragen, selecteer eerst bij welk team je hoort. Houd dit team in gedachten bij het beantwoorden van de vragen. Ik hoor bij team:"].unique().tolist()
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

    # selecting relevant team data if not all teams
    if team != "all":
        assessment = assessment[assessment["Voor het beantwoorden van de vragen, selecteer eerst bij welk team je hoort. Houd dit team in gedachten bij het beantwoorden van de vragen. Ik hoor bij team:"] == team]
    assessment = assessment.reset_index(drop=True)

    # removing " / 5"
    assessment = assessment.replace(to_replace=r"\s\/\s5", value="", regex=True)

    # storing values as integers
    assessment.loc[:, 'Ik maak me zorgen over dat ik niet aan de verwachtingen van anderen voldoe.' : 'Hoger management hecht waarde aan het inwinnen van input van alle niveaus binnen de organisatie voordat grote/belangrijke beslissingen worden genomen.'] = \
        assessment.loc[:, 'Ik maak me zorgen over dat ik niet aan de verwachtingen van anderen voldoe.' : 'Hoger management hecht waarde aan het inwinnen van input van alle niveaus binnen de organisatie voordat grote/belangrijke beslissingen worden genomen.'].astype("int")
    assessment.loc[:, 'Ik vind mijn werk persoonlijk betekenisvol en belangrijk.' : 'Ons team leeft naar en volgt een sterke set van waarden.'] = \
        assessment.loc[:, 'Ik vind mijn werk persoonlijk betekenisvol en belangrijk.' : 'Ons team leeft naar en volgt een sterke set van waarden.'].astype("int")
    assessment.loc[:, 'Teamleden zien zichzelf meer als individu dan als team.':'Teamleden voelen zich persoonlijk verantwoordelijk voor hun taken.'] = \
        assessment.loc[:, 'Teamleden zien zichzelf meer als individu dan als team.':'Teamleden voelen zich persoonlijk verantwoordelijk voor hun taken.'].astype("int")

    # reversing necessary items
    reversed_items = ["Ik krijg altijd de reacties die ik wil of verwacht.", "Ik heb het gevoel dat ik mensen bijna alles over mijzelf kan vertellen.", "Het is gemakkelijk voor mij om medeleven te tonen aan anderen.",
                    "Teamleden zien zichzelf meer als individu dan als team.", "Als je een fout maakt in dit team, wordt dit vaak tegen je gebruikt.",
                    "Mensen in dit team wijzen anderen soms af omdat ze anders zijn.", "Het is moeilijk om andere leden van dit team om hulp te vragen.",
                    "De grappen in ons team zijn niet altijd bedoeld om alleen maar de sfeer goed te houden.", "Het team maakt niet optimaal gebruik van mijn vaardigheden.",
                    "Sommige leden van het team zijn niet zo bekwaam als we nodig hebben.", "Sommige leden van het team zijn niet zo betrouwbaar als we nodig hebben."]

    for item in reversed_items:
        assessment.loc[:, item] = 6 - assessment.loc[:, item]

    # Reformulate reversed items to be aligned with reversed scoring
    reformed_items = ["Ik ben vaak verrast door de reacties die ik van mensen krijg.", "Ik ben terughoudend om mensen over mezelf te vertellen.", "Ik vind het moeilijk om medeleven te tonen aan anderen.",
                  "Leden zien zichzelf meer als een team dan als individuen.", "Als je in dit team een fout maakt, wordt die vaak vergeven.",
                  "Mensen in dit team accepteren dat andere mensen anders zijn.", "Het is gemakkelijk om andere leden van dit team om hulp te vragen.",
                  "De grappen en plagerijen in ons team zijn altijd goedbedoeld.", "Het team maakt goed gebruik van mijn vaardigheden.",
                  "Alle leden van het team zijn zo competent als we nodig hebben.", "Alle leden van het team zijn zo betrouwbaar als we nodig hebben."]

    reform_dict = dict(zip(reversed_items, reformed_items))

    assessment.rename(columns=reform_dict, inplace=True)


    # Attachment
    attachment = assessment.loc[:, "Ik maak me zorgen over dat ik niet aan de verwachtingen van anderen voldoe." : "Ik vind het moeilijk om medeleven te tonen aan anderen."]

    # Making lists of anxiety and avoidance items
    attachment_items = attachment.columns.tolist()
    anxiety = ["Ik maak me zorgen over dat ik niet aan de verwachtingen van anderen voldoe.",
            "Ik maak me zorgen over wat anderen van me denken.", "Ik maak me zorgen dat anderen mij minder waarderen dan ik hen waardeer.",
            "Ik raak geïrriteerd als mensen mij niet de steun of erkenning geven die ik denk dat ik verdien.",
            "Ik ben vaak verrast door de reacties die ik van mensen krijg.", "Soms moet ik boos worden zodat mensen mij opmerken."]
    avoidance = ['Ik begrijp niet waarom mensen soms zo emotioneel worden.',
                'Het voelt ongemakkelijk als iemand van het werk me in vertrouwen wil nemen.',
                'Ik maak me zorgen dat mensen negatief zullen reageren of me niet zullen mogen, als ze zien hoe ik echt ben.',
                'Ik werk het liefst onafhankelijk van anderen.',
                "Ik geef er de voorkeur aan om niet te vriendschappelijk met collega's om te gaan.",
                "Ik ben terughoudend om mensen over mezelf te vertellen.",
                'Ik ben terughoudend om me open te stellen voor anderen.',
                "Ik vind het moeilijk om medeleven te tonen aan anderen."]

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
    plt.xlabel("Vertrouwen")
    plt.ylabel("Assertiviteit")
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
    plt.text(2.45, 2.10, "Transformatie", fontsize=18, color="darkgray")
    plt.text(4.32, 2.10, "Dominantie", fontsize=18, color="darkgray")
    plt.text(2.25, 4.10, "Verzorgen", fontsize=18, color="darkgray")
    # plt.text(4.40, 4.10, "Controle", fontsize=18, color="darkgray", alpha=0.35)

    fig.set_facecolor("white") # set background colour to white to make labels and title visible in dark mode

    # save and clear
    plt.savefig(f"{output_dir}/Attachment_plot_{team}.png", transparent=True)
    plt.close()

    # Add mean to each item and overall score in attachment dataframe
    attachment.loc["mean"] = attachment.mean()

    # Rounding
    for col in list(attachment.columns):
        attachment.loc["mean", col] = round(attachment.loc['mean', col], 2)

    # Organisational core
    org_core = assessment.loc[:, "Onze organisatie heeft een duidelijk doel waar mensen in geloven en zich mee verbinden.":"Hoger management hecht waarde aan het inwinnen van input van alle niveaus binnen de organisatie voordat grote/belangrijke beslissingen worden genomen."]

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

    # Team core
    team_core = assessment.loc[:, "Het is voor mij duidelijk waarom dit team bestaat." : "Noem de drie belangrijkste waarden die dit team volgt en naleeft (als dat het geval is)."]

    # Calculating team core score and adding to overall df
    # creating list of variables
    team_core_items = team_core.iloc[:, 2:-2].columns.tolist()

    # adding total core score to team_core DataFrame
    team_core.loc[:, "team_core"] = team_core.loc[:, team_core_items].sum(axis=1) / len(team_core_items)

    # Round all core scores
    team_core.loc[:, "team_core"] = team_core["team_core"].round(2)

    # Set dtype of core column
    team_core.loc[:, 'team_core'] = team_core['team_core'].astype(float)

    # Add mean scores
    team_core.loc["mean"] = team_core[team_core_items].mean()

    # Add rounded mean to each column
    for col in list(team_core.columns):
        team_core.loc["mean", col] = round(team_core.loc['mean', col], 2)

    # Add core mean
    team_core.loc['mean', 'team_core'] = team_core.loc[:"mean", "team_core"].mean()

    # Add core scores to assessment df
    assessment.loc[:, "team_core"] = team_core.loc[:, 'team_core']

    # Create numeric df
    team_core_numeric = team_core.iloc[:-1, :-1].select_dtypes(include='number')

    # Variance
    team_core_sorted_var = team_core_numeric.var().sort_values(ascending=False).round(2)
    team_core_top_var = team_core_sorted_var[:3]
    team_core_bottom_var = team_core_sorted_var[-3:]    

    # # Variance for each item in team core
    # team_core_variance = {}
    # for col in team_core_items:
    #     team_core_variance[col] = round(team_core.drop('mean', axis=0)[col].var(), 2)

    # team_core_variance = {k: v for k, v in sorted(team_core_variance.items(), key=lambda item: item[1])}

    # # 3 lowest variance items
    # team_core_low_list = list(team_core_variance.keys())[:3]
    # team_core_low_var = team_core.loc['mean', team_core_low_list]

    # # 3 highest variance items
    # team_core_high_list = list(team_core_variance.keys())[:-3]
    # team_core_high_var = team_core.loc['mean', team_core_high_list]

    # Write team core stats to text file
    with open(f'{output_dir}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
        # Descriptive stats
        f.write(f'\n \nTeam core: \n{round(team_core.loc[:"mean", "team_core"].describe()[["mean", "min", "max"]], 2)}')
        # Lowest and highest scoring items
        f.write(f'\n \nThree lowest scoring team_core items: \n{team_core.iloc[-1, 2:-2].sort_values()[:3]}\n \nThree highest scoring team_core items: \n{team_core.iloc[-1, 2:-2].sort_values(ascending=False)[:3]}')
        ## Lowest and highest variance items
        f.write(f'\n \nThree lowest variance team_core items: \n{team_core_bottom_var} \n \nThree highest variance team_core items: \n{team_core_top_var}')
        # It is clear to me why this team exists
        f.write('\n \nIt is clear to me why this team exists')
        for i in range(len(team_core["Het is voor mij duidelijk waarom dit team bestaat."]) - 1):
            f.write(f'\n{team_core.loc[i, "Het is voor mij duidelijk waarom dit team bestaat."]}')
        # Waarom bestaat dit team?
        f.write('\n \nWWaarom bestaat dit team?')
        for i in range(len(team_core["Waarom bestaat dit team?"]) - 1):
            f.write(f'\n{team_core.loc[i, "Waarom bestaat dit team?"]}')
        # Noem de drie belangrijkste waarden die dit team volgt en naleeft (als dat het geval is).
        f.write('\n \nNoem de drie belangrijkste waarden die dit team volgt en naleeft (als dat het geval is).')
        for i in range(len(team_core["Noem de drie belangrijkste waarden die dit team volgt en naleeft (als dat het geval is)."]) - 1):
            f.write(f'\n{team_core.loc[i, "Noem de drie belangrijkste waarden die dit team volgt en naleeft (als dat het geval is)."]}')

    # Team safety
    safety = assessment.loc[:, "Leden zien zichzelf meer als een team dan als individuen.":"De grappen en plagerijen in ons team zijn altijd goedbedoeld."]

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
    with open(f'{output_dir}/Assessment results_{team}.txt', 'a') as f:
        # Descriptive stats
        f.write(f'\n \nTeam safety: \n{round(safety.loc[:"mean", "safety"].describe()[["mean", "min", "max"]], 2)}')
        # Lowest and highest scoring items
        f.write(f'\n \nThree lowest scoring safety items: \n{safety.iloc[-1, :-1].sort_values()[:3]}\n \nThree highest scoring safety items: \n{safety.iloc[-1, :-1].sort_values(ascending=False)[:3]}')
        ## Lowest and highest variance items
        f.write(f'\n \nThree lowest variance safety items: \n{safety_bottom_var} \n \nThree highest variance safety items: \n{safety_top_var}')

    # Team dependability
    dependability = assessment.loc[:, "Rollen en verantwoordelijkheden zijn duidelijk toebedeeld in het team.":"Teamleden voelen zich persoonlijk verantwoordelijk voor hun taken."]

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

    # team_core
    team_core_dir = f"{output_dir}/team_core"
    os.makedirs(team_core_dir, exist_ok=True)

    scale_plots("team_core", team, team_core_dir, color="gray")

    # ## Highest scoring team_core items
    # bar_plots(team_core.iloc[-1, 2:-2].sort_values(ascending=False)[:3].index.tolist(),
    #                  team_core, "gray", team_core_dir, team)

    # ## Lowest scoring team_core items
    # bar_plots(team_core.iloc[-1, 2:-2].sort_values()[:3].index.tolist(),
    #                     team_core, "gray", team_core_dir, team)

    # # Highest variance team_core items
    # bar_plots(team_core_high_list, team_core, "gray", team_core_dir, team)

    # # Lowest variance team_core items
    # bar_plots(team_core_low_list, team_core, "gray", team_core_dir, team)

    # Plot all numeric team_core items
    bar_plots(team_core_numeric, "gray", team_core_dir, team)
    ranked_bar_plot(team_core_numeric, "team_core", team, team_core_dir)
    ranked_variance_plot(team_core_numeric, "team_core", team, team_core_dir)

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
    scales = ["org_core", "team_core", "safety", "dependability"]
    plt.figure(figsize=(10, 6))
    plt.bar(assessment[scales].var().index, assessment[scales].var(), color="gray")
    plt.xticks(ticks=range(len(scales)), labels=["Organisational core", "Team core", "Psychological safety", "Team dependability"])
    plt.tight_layout()
    plt.savefig(f'{output_dir}/variance_{team}.png', transparent=True)
    plt.close()

    # Plot scores for all scales
    plt.figure(figsize=(10, 6))
    plt.bar(assessment[scales].mean().index, assessment.loc["mean", scales], color="gray")
    plt.yticks(ticks=np.arange(1, 6, 1))
    plt.xticks(ticks=range(len(scales)), labels=["Organisational core", "Team core", "Psychological safety", "Team dependability"])
    plt.tight_layout()
    plt.savefig(f'{output_dir}/scale_scores_{team}.png', transparent=True)

print("Analysis complete. Results exported to Excel and text files.")