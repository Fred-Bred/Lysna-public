import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
import re
import numpy as np
import sys
import os
import textwrap
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import time

from lysna import language
from lysna.plotting import *
from lysna.language import English, Dutch, Danish


# Functions
def browse_file():
    file_path = filedialog.askopenfilename(title="Select a file", filetypes=[("All files", "*.*")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

def browse_output_directory():
    directory_path = filedialog.askdirectory(title="Select output directory")
    if directory_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, directory_path)

def run_analysis():
    try:
        file_name = file_entry.get()
        plots = plot_var.get()
        dynamic_plots = dynamic_plots_var.get()
        master_dir = output_entry.get()
        language_choice = language_var.get()

        if language_choice == "English":
            lang = English()
        elif language_choice == "Danish":
            lang = Danish()
        elif language_choice == "Dutch":
            lang = Dutch()

        if not file_name:
            messagebox.showerror("Error", "Please select a file.")
            return
        
        if not file_name.endswith(('.csv', '.xlsx')):  
            messagebox.showerror("Error", "Please select a valid CSV or Excel file.")
            return
        
        if not master_dir:
            messagebox.showerror("Error", "Please select an output directory.")
            return
        
        # Read the file
        try:
            # path = file_name + ".csv"
            raw_assessment = pd.read_csv(file_name)
        except:
            try:
                # path = file_name + ".xlsx"
                raw_assessment = pd.read_excel(file_name)
            except:
                messagebox.ERROR("File not found")


        # Data cleaning
        # dropping useless data
        assessment = raw_assessment.drop(["Date Time", "User", "Page submitted"], axis=1)
        teams = ["all"]
        try:
            assessment = assessment.dropna(subset=lang.team_filter)
            mode = "multiple teams"
            teams = teams + raw_assessment[lang.team_filter].unique().tolist()
        except:
            mode = "single team"

        assessment = assessment.reset_index(drop=True)

        analysis_type_label.config(text=f"Running analysis for: {mode}")
        root.update_idletasks()

        # Make output folders for multiple teams
        if mode == "multiple teams":
            for team in teams:
                os.makedirs("/".join([master_dir, team]), exist_ok=True)
        else:
            os.makedirs("/".join([master_dir, "Results"]), exist_ok=True)

        # Run analysis for each unique team
        for t in teams:
            # Data cleaning
            # dropping useless data
            team = t
            assessment = raw_assessment.drop(["Date Time", "User", "Page submitted"], axis=1)

            # Specify output dir
            if mode == "multiple teams":
                output_dir = f"{master_dir}/{team}"
            else:
                output_dir = f"{master_dir}/Results"

            # selecting relevant team data if not all teams
            if team != "all":
                assessment = assessment[assessment[lang.team_filter] == team]
            assessment = assessment.reset_index(drop=True)

            # removing " / 5"
            assessment = assessment.replace(to_replace=r"\s\/\s5", value="", regex=True)

            # storing values as integers
            assessment.loc[:, lang.numeric_idxs[0][0] : lang.numeric_idxs[0][1]] = \
                assessment.loc[:, lang.numeric_idxs[0][0] : lang.numeric_idxs[0][1]].astype("int")
            assessment.loc[:, lang.numeric_idxs[1][0] : lang.numeric_idxs[1][1]] = \
                assessment.loc[:, lang.numeric_idxs[1][0] : lang.numeric_idxs[1][1]].astype("int")
            assessment.loc[:, lang.numeric_idxs[2][0] : lang.numeric_idxs[2][1]] = \
                assessment.loc[:, lang.numeric_idxs[2][0] : lang.numeric_idxs[2][1]].astype("int")

            # reversing necessary items
            reversed_items = lang.reversed_items

            for item in reversed_items:
                assessment.loc[:, item] = 6 - assessment.loc[:, item]

            # Reformulate reversed items to be aligned with reversed scoring
            reformed_items = lang.reformed_items

            reform_dict = dict(zip(reversed_items, reformed_items))

            assessment.rename(columns=reform_dict, inplace=True)


            # Attachment
            attachment = assessment.loc[:, lang.attachment_idxs[0] : lang.attachment_idxs[1]]

            # Making lists of anxiety and avoidance items
            attachment_items = attachment.columns.tolist()
            anxiety = lang.anxiety
            
            avoidance = lang.avoidance

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
            plt.xlabel(lang.labels["Trust"])
            plt.ylabel(lang.labels["Confidence"])
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
            plt.text(2.45, 2.10, lang.labels["Transformation"], fontsize=18, color="darkgray")
            plt.text(4.35, 2.10, lang.labels["Dominance"], fontsize=18, color="darkgray")
            plt.text(2.25, 4.10, lang.labels["Nurture"], fontsize=18, color="darkgray")
            # plt.text(4.42, 4.10, "Fear/dominance", fontsize=18, color="darkgray", alpha=0.35)

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
            org_core = assessment.loc[:, lang.org_core_idxs[0] : lang.org_core_idxs[1]]

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

            # Write organisational core stats to text file
            with open(f'{output_dir}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
                # Descriptive stats
                f.write(f'\n \nOrganisational core: \n{round(org_core.loc[:"mean", "org_core"].describe()[["mean", "min", "max"]], 2)}')
                # Lowest and highest scoring items
                f.write(f'\n \nThree lowest scoring org_core items: \n{org_core.iloc[-1, :-1].sort_values()[:3]}\n \nThree highest scoring org_core items: \n{org_core.iloc[-1, :-1].sort_values(ascending=False)[:3]}')
                ## Lowest and highest variance items
                f.write(f'\n \nThree lowest variance org_core items: \n{org_core_bottom_var} \n \nThree highest variance org_core items: \n{org_core_top_var}')

            # Team core
            team_core = assessment.loc[:, lang.team_core_idxs[0] : lang.team_core_idxs[1]]

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

            # Write team core stats to text file
            with open(f'{output_dir}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
                # Descriptive stats
                f.write(f'\n \nTeam core: \n{round(team_core.loc[:"mean", "team_core"].describe()[["mean", "min", "max"]], 2)}')
                # Lowest and highest scoring items
                f.write(f'\n \nThree lowest scoring team_core items: \n{team_core.iloc[-1, 2:-2].sort_values()[:3]}\n \nThree highest scoring team_core items: \n{team_core.iloc[-1, 2:-2].sort_values(ascending=False)[:3]}')
                ## Lowest and highest variance items
                f.write(f'\n \nThree lowest variance team_core items: \n{team_core_bottom_var} \n \nThree highest variance team_core items: \n{team_core_top_var}')
                
                # Free form items 1-3
                f.write(f'\n \n{lang.free_form[0]}')
                for i in range(len(team_core[lang.free_form[0]]) - 1):
                    f.write(f'\n{team_core.loc[i, lang.free_form[0]]}')
                
                f.write(f'\n \n{lang.free_form[1]}')
                for i in range(len(team_core[lang.free_form[1]]) - 1):
                    f.write(f'\n{team_core.loc[i, lang.free_form[1]]}')
                
                f.write(f'\n \n{lang.free_form[2]}')
                for i in range(len(team_core[lang.free_form[2]]) - 1):
                    f.write(f'\n{team_core.loc[i, lang.free_form[2]]}')

            # Team safety
            safety = assessment.loc[:, lang.safety_idxs[0] : lang.safety_idxs[1]]

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

            # Write safety stats to text file
            with open(f'{output_dir}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
                # Descriptive stats
                f.write(f'\n \nTeam safety: \n{round(safety.loc[:"mean", "safety"].describe()[["mean", "min", "max"]], 2)}')
                # Lowest and highest scoring items
                f.write(f'\n \nThree lowest scoring safety items: \n{safety.iloc[-1, :-1].sort_values()[:3]}\n \nThree highest scoring safety items: \n{safety.iloc[-1, :-1].sort_values(ascending=False)[:3]}')
                ## Lowest and highest variance items
                f.write(f'\n \nThree lowest variance safety items: \n{safety_bottom_var} \n \nThree highest variance safety items: \n{safety_top_var}')

            # Team dependability
            dependability = assessment.loc[:, lang.dependability_idxs[0] : lang.dependability_idxs[1]]

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
            if plots == False:
                continue # End this loop iteration if plots are not requested

            print(f"Plotting for {team}...")

            # # org_core
            org_core_dir = f"{output_dir}/org_core"
            os.makedirs(org_core_dir, exist_ok=True)

            scale_plots(assessment, "org_core", team, org_core_dir, color="gray")

            # Plot all numeric org_core items
            bar_plots(org_core_numeric, "gray", org_core_dir, team, dynamic=dynamic_plots)
            ranked_bar_plot(org_core_numeric, "org_core", team, org_core_dir, dynamic=dynamic_plots)
            ranked_variance_plot(org_core_numeric, "org_core", team, org_core_dir)

            # team_core
            team_core_dir = f"{output_dir}/team_core"
            os.makedirs(team_core_dir, exist_ok=True)

            scale_plots(assessment, "team_core", team, team_core_dir, color="gray", dynamic=dynamic_plots)

            # Plot all numeric team_core items
            bar_plots(team_core_numeric, "gray", team_core_dir, team, dynamic=dynamic_plots)
            ranked_bar_plot(team_core_numeric, "team_core", team, team_core_dir, dynamic=dynamic_plots)
            ranked_variance_plot(team_core_numeric, "team_core", team, team_core_dir)

            # safety
            safety_dir = f"{output_dir}/safety"
            os.makedirs(safety_dir, exist_ok=True)

            scale_plots(assessment, "safety", team, safety_dir, color="gray", dynamic=dynamic_plots)

            # Plot all numeric safety items
            bar_plots(safety_numeric, "gray", safety_dir, team, dynamic=dynamic_plots)
            ranked_bar_plot(safety_numeric, "safety", team, safety_dir, dynamic=dynamic_plots)
            ranked_variance_plot(safety_numeric, "safety", team, safety_dir)

            # # dependability
            dependability_dir = f"{output_dir}/dependability"
            os.makedirs(dependability_dir, exist_ok=True)

            scale_plots(assessment, "dependability", team, dependability_dir, color="gray", dynamic=dynamic_plots)

            # Plot all numeric dependability items
            bar_plots(dependability_numeric, "gray", dependability_dir, team, dynamic=dynamic_plots)
            ranked_bar_plot(dependability_numeric, "dependability", team, dependability_dir, dynamic=dynamic_plots)
            ranked_variance_plot(dependability_numeric, "dependability", team, dependability_dir)

            # Plot variance for all scales
            scales = ["org_core", "team_core", "safety", "dependability"]
            plt.figure(figsize=(10, 6))
            plt.bar(assessment[scales].var().index, assessment[scales].var(), color="gray")
            plt.xticks(ticks=range(len(scales)), labels=list(lang.scales.values()))
            plt.tight_layout()
            plt.savefig(f'{output_dir}/variance_{team}.png', transparent=True)
            plt.close()

            # Plot scores for all scales with conditional coloring
            plt.figure(figsize=(10, 6))
            values = assessment.loc["mean", scales]
            colors = get_bar_colors(values) if dynamic_plots else "gray"
            plt.bar(range(len(scales)), values, color=colors)
            plt.yticks(ticks=np.arange(1, 6, 1))
            plt.xticks(ticks=range(len(scales)), labels=list(lang.scales.values()))
            plt.tight_layout()
            plt.savefig(f'{output_dir}/scale_scores_{team}.png', transparent=True)
            plt.close()


        messagebox.showinfo("Success", "Analysis completed!")
        # Update the result label
        result_label.config(text="Analysis completed successfully! Output saved.", fg="green")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        # Update the result label
        result_label.config(text=f"Error: {e}", fg="red")

if __name__ == "__main__":
    # Create the main window
    root = tk.Tk()
    root.title("Lysna Assessment Analysis Tool")
    # root.geometry("800x400")

    # File name entry
    tk.Label(root, text="Selected file:").grid(row=0, column=0, padx=10, pady=10)
    file_entry = tk.Entry(root, width=50)
    file_entry.grid(row=0, column=1, padx=10, pady=10)

    # Browse button
    browse_button = tk.Button(root, text="Browse", command=browse_file)
    browse_button.grid(row=0, column=2, padx=10, pady=10)

    # Output directory entry
    tk.Label(root, text="Output directory:").grid(row=1, column=0, padx=10, pady=10)
    output_entry = tk.Entry(root, width=50)
    output_entry.grid(row=1, column=1, padx=10, pady=10)

    # Browse button for output directory
    browse_output_button = tk.Button(root, text="Browse", command=browse_output_directory)
    browse_output_button.grid(row=1, column=2, padx=10, pady=10)

    # Language selection
    tk.Label(root, text="Select language:").grid(row=2, column=0, padx=10, pady=10)
    language_var = tk.StringVar()
    language_combobox = ttk.Combobox(root, textvariable=language_var, values=["English", "Danish", "Dutch"])
    language_combobox.grid(row=2, column=1, padx=10, pady=10)
    language_combobox.current(0)  # Set default value to English

    # Produce plots checkbox
    plot_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Produce plots", variable=plot_var).grid(row=3, column=0, columnspan=2, pady=10)
    
    # Dynamic plots checkbox
    dynamic_plots_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Dynamic plots", variable=dynamic_plots_var).grid(row=4, column=0, columnspan=2, pady=10)

    # Run button
    run_button = tk.Button(root, text="Run Analysis", command=run_analysis)
    run_button.grid(row=5, column=0, columnspan=2, pady=10)

    # Analysis type label
    analysis_type_label = tk.Label(root, text="", fg="blue")
    analysis_type_label.grid(row=4, column=0, columnspan=3, pady=10)

    # Result label
    result_label = tk.Label(root, text="", fg="green")
    result_label.grid(row=5, column=0, columnspan=3, pady=10)


    # Start the GUI event loop
    root.mainloop()