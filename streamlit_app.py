import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
import tempfile
import time
import shutil
from io import BytesIO
import zipfile

from lysna.language import English, Danish, Dutch
from lysna.plotting import *


def run_analysis(uploaded_file, output_dir, language_choice, plots, dynamic_plots):
    """
    Run the analysis with the provided parameters.
    Returns success status and message.
    """
    try:
        if language_choice == "English":
            lang = English()
        elif language_choice == "Danish":
            lang = Danish()
        elif language_choice == "Dutch":
            lang = Dutch()

        # Read the file
        try:
            if uploaded_file.name.endswith('.csv'):
                raw_assessment = pd.read_csv(uploaded_file)
            elif uploaded_file.name.endswith('.xlsx'):
                raw_assessment = pd.read_excel(uploaded_file)
            else:
                return False, "Please select a valid CSV or Excel file."
        except Exception as e:
            return False, f"Error reading file: {str(e)}"

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

        # Progress indicator with metrics and activity feed
        st.subheader("Analysis Progress")
        
        # Top-level metrics
        col1, col2, col3 = st.columns(3)
        with col1:
            progress_metric = st.metric("Progress", "0%")
        with col2:
            current_team_metric = st.metric("Current Team", "Starting...")
        with col3:
            eta_metric = st.metric("Status", "🚀 Initializing")
        
        # Activity feed
        st.markdown("**📋 Activity Log**")
        activity_container = st.empty()
        activity_messages = []
        
        def update_activity(message, level="info"):
            icons = {"info": "ℹ️", "success": "✅", "warning": "⚠️", "processing": "⚙️", "plot": "📈"}
            timestamp = time.strftime("%H:%M:%S")
            activity_messages.append(f"`{timestamp}` {icons.get(level, 'ℹ️')} {message}")
            
            # Update the container with latest messages (limit to 6 and reverse order)
            with activity_container.container():
                for msg in reversed(activity_messages[-6:]):
                    st.markdown(msg)
        
        update_activity("Analysis started")
        update_activity(f"Detected {mode} with {len(teams)} teams")

        # Make output folders for multiple teams
        if mode == "multiple teams":
            for team in teams:
                os.makedirs("/".join([output_dir, team]), exist_ok=True)
        else:
            os.makedirs("/".join([output_dir, "Results"]), exist_ok=True)

        total_teams = len(teams)
        start_time = time.time()
        
        # Run analysis for each unique team
        for i, t in enumerate(teams):
            # Update metrics
            progress_pct = int((i / total_teams) * 100)
            progress_metric.metric("Progress", f"{progress_pct}%")
            current_team_metric.metric("Current Team", t)
            eta_metric.metric("Status", "⚙️ Processing")
            
            update_activity(f"Processing team: {t} ({i+1}/{total_teams})", "processing")
            
            # Data cleaning
            # dropping useless data
            team = t
            assessment = raw_assessment.drop(["Date Time", "User", "Page submitted"], axis=1)

            # Specify output dir
            if mode == "multiple teams":
                output_path = f"{output_dir}/{team}"
            else:
                output_path = f"{output_dir}/Results"

            # selecting relevant team data if not all teams
            if team != "all":
                assessment = assessment[assessment[lang.team_filter] == team]
            assessment = assessment.reset_index(drop=True)

            # removing " / 5"
            assessment = assessment.replace(to_replace=r"\s\/\s5", value="", regex=True)

            update_activity(f"Data cleaning completed for {t}", "success")

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
            for j, score in enumerate(attachment.loc[:, "anxiety"]):
                attachment.loc[j, "anxiety"] = round(attachment.loc[j, "anxiety"], 2) # type: ignore

            for j, score in enumerate(attachment.loc[:, 'avoidance']):
                attachment.loc[j, 'avoidance'] = round(attachment.loc[j, 'avoidance'], 2) # type: ignore

            # Round scores after old code broke and add to assessment df
            for j, score in enumerate(assessment.loc[:, "anxiety"]):
                assessment.loc[j, "anxiety"] = round(assessment.loc[j, "anxiety"], 2) # type: ignore

            for j, score in enumerate(assessment.loc[:, 'avoidance']):
                assessment.loc[j, 'avoidance'] = round(assessment.loc[j, 'avoidance'], 2) # type: ignore

            # Set dtypes to float
            attachment.loc[:, 'anxiety'] = attachment['anxiety'].astype(float)
            attachment.loc[:, 'avoidance'] = attachment['avoidance'].astype(float)
            assessment.loc[:, 'anxiety'] = assessment['anxiety'].astype(float)
            assessment.loc[:, 'avoidance'] = assessment['avoidance'].astype(float)

            # Anxiety descriptive statistics
            round(assessment.loc[:"mean", "anxiety"].describe()[["mean", "min", "max"]], 2)

            # Store anxiety scores in text file
            with open(f'{output_path}/Assessment results_{team}.txt', 'w', encoding='utf-8') as f:
                f.write(f'Team anxiety scores: \n{round(assessment.loc[:"mean", "anxiety"].describe()[["mean", "min", "max"]], 2)}')

            # Avoidance descriptive statistics
            round(assessment.loc[:"mean", "avoidance"].describe()[["mean", "min", "max"]], 2)

            # Store avoidance scores in text file
            with open(f'{output_path}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
                f.write(f'\n \nTeam avoidance scores: \n{round(assessment.loc[:"mean", "avoidance"].describe()[["mean", "min", "max"]], 2)}')

            # Plot attachment scores
            fig, ax = plt.subplots(figsize=(10, 7), dpi=400)

            x = attachment['avoidance']
            y = attachment['anxiety']
            markers = ["d", "v", "s", "*", "^", "*", "^", "x", "p", "P", "D", "H", "1", "2", "3", "4", "<", ">"]
            for xp, yp, m in zip(x, y, markers):
                plt.scatter(xp, yp, marker=m, s=50, alpha=0.75)

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
            plt.savefig(f"{output_path}/Attachment_plot_{team}.png", transparent=True)
            plt.close()

            # Add mean to each item and overall score in attachment dataframe
            attachment.loc["mean"] = attachment.mean()

            # Rounding
            for col in list(attachment.columns):
                attachment.loc["mean", col] = round(attachment.loc['mean', col], 2) # type: ignore

            update_activity(f"Attachment analysis completed for {t}", "success")

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
                org_core.loc["mean", col] = round(org_core.loc['mean', col], 2) # type: ignore

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
            with open(f'{output_path}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
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
                team_core.loc["mean", col] = round(team_core.loc['mean', col], 2) # type: ignore

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
            with open(f'{output_path}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
                # Descriptive stats
                f.write(f'\n \nTeam core: \n{round(team_core.loc[:"mean", "team_core"].describe()[["mean", "min", "max"]], 2)}')
                # Lowest and highest scoring items
                f.write(f'\n \nThree lowest scoring team_core items: \n{team_core.iloc[-1, 2:-2].sort_values()[:3]}\n \nThree highest scoring team_core items: \n{team_core.iloc[-1, 2:-2].sort_values(ascending=False)[:3]}')
                ## Lowest and highest variance items
                f.write(f'\n \nThree lowest variance team_core items: \n{team_core_bottom_var} \n \nThree highest variance team_core items: \n{team_core_top_var}')
                
                # Free form items 1-3
                f.write(f'\n \n{lang.free_form[0]}')
                for j in range(len(team_core[lang.free_form[0]]) - 1):
                    f.write(f'\n{team_core.loc[j, lang.free_form[0]]}')
                
                f.write(f'\n \n{lang.free_form[1]}')
                for j in range(len(team_core[lang.free_form[1]]) - 1):
                    f.write(f'\n{team_core.loc[j, lang.free_form[1]]}')
                
                f.write(f'\n \n{lang.free_form[2]}')
                for j in range(len(team_core[lang.free_form[2]]) - 1):
                    f.write(f'\n{team_core.loc[j, lang.free_form[2]]}')

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
                safety.loc["mean", col] = round(safety.loc['mean', col], 2) # type: ignore

            # adding total core score to assessment dataframe
            assessment.loc[:, "safety"] = safety.loc[:, 'safety']

            # Create numeric df
            safety_numeric = safety.iloc[:-1, :-1].select_dtypes(include='number')

            # Variance
            safety_sorted_var = safety_numeric.var().sort_values(ascending=False).round(2)
            safety_top_var = safety_sorted_var[:3]
            safety_bottom_var = safety_sorted_var[-3:]

            # Write safety stats to text file
            with open(f'{output_path}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
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
                dependability.loc["mean", col] = round(dependability.loc['mean', col], 2) # type: ignore

            # adding total dependability score to assessment dataframe
            assessment.loc[:, "dependability"] = dependability.loc[:, 'dependability']

            # Create numeric df
            dependability_numeric = dependability.iloc[:-1, :-1].select_dtypes(include='number')

            # Variance
            dependability_sorted_var = dependability_numeric.var().sort_values(ascending=False).round(2)
            dependability_top_var = dependability_sorted_var[:3]
            dependability_bottom_var = dependability_sorted_var[-3:]

            # Write dependability stats to text file
            with open(f'{output_path}/Assessment results_{team}.txt', 'a', encoding='utf-8') as f:
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
            assessment.to_excel(f"{output_path}/Team assessment results {team}.xlsx")
            
            update_activity(f"Excel file generated for {t}", "success")

            ### PLOTTING (optional)
            if not plots:
                update_activity(f"Completed team: {t} (no plots)", "success")
                # Estimate time remaining
                elapsed = time.time() - start_time
                if i > 0:
                    avg_time_per_team = elapsed / (i + 1)
                    remaining_teams = total_teams - (i + 1)
                    if remaining_teams > 0:
                        eta_seconds = remaining_teams * avg_time_per_team
                        eta_text = f"~{int(eta_seconds)}s remaining"
                        eta_metric.metric("ETA", eta_text)
                continue # End this loop iteration if plots are not requested

            update_activity(f"Generating plots for {t}", "plot")

            # # org_core
            org_core_dir = f"{output_path}/org_core"
            os.makedirs(org_core_dir, exist_ok=True)

            scale_plots(assessment, "org_core", team, org_core_dir, color="gray")

            # Plot all numeric org_core items
            bar_plots(org_core_numeric, "gray", org_core_dir, team, dynamic=dynamic_plots)
            ranked_bar_plot(org_core_numeric, "org_core", team, org_core_dir, dynamic=dynamic_plots)
            ranked_variance_plot(org_core_numeric, "org_core", team, org_core_dir)

            # team_core
            team_core_dir = f"{output_path}/team_core"
            os.makedirs(team_core_dir, exist_ok=True)

            scale_plots(assessment, "team_core", team, team_core_dir, color="gray", dynamic=dynamic_plots)

            # Plot all numeric team_core items
            bar_plots(team_core_numeric, "gray", team_core_dir, team, dynamic=dynamic_plots)
            ranked_bar_plot(team_core_numeric, "team_core", team, team_core_dir, dynamic=dynamic_plots)
            ranked_variance_plot(team_core_numeric, "team_core", team, team_core_dir)

            # safety
            safety_dir = f"{output_path}/safety"
            os.makedirs(safety_dir, exist_ok=True)

            scale_plots(assessment, "safety", team, safety_dir, color="gray", dynamic=dynamic_plots)

            # Plot all numeric safety items
            bar_plots(safety_numeric, "gray", safety_dir, team, dynamic=dynamic_plots)
            ranked_bar_plot(safety_numeric, "safety", team, safety_dir, dynamic=dynamic_plots)
            ranked_variance_plot(safety_numeric, "safety", team, safety_dir)

            # # dependability
            dependability_dir = f"{output_path}/dependability"
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
            plt.savefig(f'{output_path}/variance_{team}.png', transparent=True)
            plt.close()

            # Plot scores for all scales with conditional coloring
            plt.figure(figsize=(10, 6))
            values = assessment.loc["mean"][scales]
            colors = get_bar_colors(values) if dynamic_plots else "gray"
            plt.bar(range(len(scales)), values, color=colors)
            plt.yticks(ticks=np.arange(1, 6, 1))
            plt.xticks(ticks=range(len(scales)), labels=list(lang.scales.values()))
            plt.tight_layout()
            plt.savefig(f'{output_path}/scale_scores_{team}.png', transparent=True)
            plt.close()
            
            update_activity(f"Completed team: {t}", "success")
            
            # Estimate time remaining
            elapsed = time.time() - start_time
            if i < total_teams - 1:  # Not the last team
                avg_time_per_team = elapsed / (i + 1)
                remaining_teams = total_teams - (i + 1)
                eta_seconds = remaining_teams * avg_time_per_team
                eta_text = f"~{int(eta_seconds)}s remaining"
                eta_metric.metric("ETA", eta_text)

        # Final status updates
        progress_metric.metric("Progress", "100%")
        current_team_metric.metric("Current Team", "All Complete")
        eta_metric.metric("Status", "✅ Completed")
        update_activity("Analysis completed successfully!", "success")
        
        return True, "Analysis completed successfully! All files have been generated."
    
    except Exception as e:
        return False, f"An error occurred: {str(e)}"


def create_download_zip(output_dir):
    """Create a ZIP file of all output files for download."""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arc_name = os.path.relpath(file_path, output_dir)
                zip_file.write(file_path, arc_name)
    zip_buffer.seek(0)
    return zip_buffer


def main():
    st.set_page_config(
        page_title="Lysna Assessment Analysis Tool",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("Lysna Assessment Analysis Tool")
    st.markdown("---")
    
    # Create sidebar for inputs
    st.sidebar.header("Configuration")
    
    # File upload
    uploaded_file = st.sidebar.file_uploader(
        "Upload your assessment file",
        type=['csv', 'xlsx'],
        help="Select a CSV or Excel file containing assessment data"
    )
    
    # Language selection
    language_choice = st.sidebar.selectbox(
        "Select language",
        options=["English", "Danish", "Dutch"],
        index=0
    )
    
    # Analysis options
    st.sidebar.subheader("Analysis Options")
    produce_plots = st.sidebar.checkbox("Produce plots", value=True)
    dynamic_plots = st.sidebar.checkbox("Dynamic plots", value=True)
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("Instructions")
        st.markdown("""
        1. **Upload File**: Select your CSV or Excel file containing assessment data
        2. **Choose Language**: Select the appropriate language for analysis to match your data
        3. **Configure Options**: 
           - Check "Produce plots" to generate visualisations
           - Check "Dynamic plots" for enhanced plot colouring
        4. **Run Analysis**: Click the button below to start processing
        5. **Download Results**: Once complete, download the ZIP file with all results
        """)
    
    with col2:
        st.header("File Information")
        if uploaded_file is not None:
            st.success(f"✅ File uploaded: {uploaded_file.name}")
            st.info(f"File size: {uploaded_file.size:,} bytes")
            
            # Show file preview
            try:
                if uploaded_file.name.endswith('.csv'):
                    df_preview = pd.read_csv(uploaded_file, nrows=5)
                else:
                    df_preview = pd.read_excel(uploaded_file, nrows=5)
                
                st.subheader("Data Preview")
                st.dataframe(df_preview, use_container_width=True)
                
                # Reset file pointer for analysis
                uploaded_file.seek(0)
            except Exception as e:
                st.error(f"Error previewing file: {str(e)}")
        else:
            st.warning("Please upload a file to continue")
    
    st.markdown("---")
    
    # Analysis section
    if uploaded_file is not None:
        st.header("Run Analysis")
        
        if st.button("Start Analysis", type="primary", use_container_width=True):
            # Create temporary directory for output
            temp_dir = tempfile.mkdtemp()
            
            # Run analysis
            success, message = run_analysis(
                uploaded_file, 
                temp_dir, 
                language_choice, 
                produce_plots, 
                dynamic_plots
            )
            
            if success:
                # Store results in session state to persist across reruns
                st.session_state['analysis_complete'] = True
                st.session_state['temp_dir'] = temp_dir
                st.session_state['language_choice'] = language_choice
                st.session_state['success_message'] = message
                st.rerun()
        
        # Display results if analysis is complete
        if st.session_state.get('analysis_complete', False):
            temp_dir = st.session_state['temp_dir']
            language_choice = st.session_state['language_choice']
            message = st.session_state['success_message']
            
            st.success(message)
            
            # Show summary of generated files
            st.subheader("Generated Files Summary")
            file_count = 0
            all_files = []
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_path, temp_dir)
                    all_files.append((rel_path, file_path))
                    file_count += 1
            
            st.info(f"✅ Analysis complete! Generated {file_count} files ready for download.")
            
            # File preview section
            st.subheader("📁 File Browser & Preview")
            
            # Group files by type/folder
            file_groups = {}
            for rel_path, full_path in all_files:
                if '/' in rel_path:
                    folder = rel_path.split('/')[0]
                    if folder not in file_groups:
                        file_groups[folder] = []
                    file_groups[folder].append((rel_path, full_path))
                else:
                    if 'Root Files' not in file_groups:
                        file_groups['Root Files'] = []
                    file_groups['Root Files'].append((rel_path, full_path))
            
            # Create expandable sections for each folder
            for folder_name, files in file_groups.items():
                with st.expander(f"📂 {folder_name} ({len(files)} files)", expanded=False):
                    for rel_path, full_path in files:
                        col1, col2 = st.columns([3, 1])
                        
                        with col1:
                            st.text(f"📄 {os.path.basename(rel_path)}")
                        
                        with col2:
                            # Add preview button for certain file types
                            file_ext = os.path.splitext(rel_path)[1].lower()
                            if file_ext in ['.txt', '.xlsx', '.png']:
                                if st.button(f"Preview", key=f"preview_{rel_path}", help=f"Preview {os.path.basename(rel_path)}"):
                                    st.session_state[f'show_preview_{rel_path}'] = True
                        
                        # Show preview if button was clicked
                        if st.session_state.get(f'show_preview_{rel_path}', False):
                            try:
                                if file_ext == '.txt':
                                    with open(full_path, 'r', encoding='utf-8') as f:
                                        content = f.read()
                                    st.text_area(f"Preview: {os.path.basename(rel_path)}", content, height=200, key=f"content_{rel_path}")
                                
                                elif file_ext == '.xlsx':
                                    try:
                                        df = pd.read_excel(full_path)
                                        st.write(f"**Preview: {os.path.basename(rel_path)}** (First 10 rows)")
                                        
                                        # Handle mixed data types that cause Arrow conversion issues
                                        df_preview = df.head(10).copy()
                                        
                                        # Convert problematic columns to string to avoid Arrow conversion errors
                                        for col in df_preview.columns:
                                            if df_preview[col].dtype == 'object':
                                                df_preview[col] = df_preview[col].astype(str)
                                        
                                        st.dataframe(df_preview, use_container_width=True)
                                        st.info(f"Full file contains {len(df)} rows and {len(df.columns)} columns")
                                    except Exception as e:
                                        # Fallback: show basic file info if preview fails
                                        st.warning(f"Cannot preview Excel file: {str(e)}")
                                        st.info(f"Excel file: {os.path.basename(rel_path)} (Preview unavailable)")
                                
                                elif file_ext == '.png':
                                    st.write(f"**Preview: {os.path.basename(rel_path)}**")
                                    st.image(full_path, caption=os.path.basename(rel_path), use_container_width=True)
                                
                                # Add close button
                                if st.button(f"❌ Close Preview", key=f"close_{rel_path}"):
                                    st.session_state[f'show_preview_{rel_path}'] = False
                                    st.rerun()
                                    
                            except Exception as e:
                                st.error(f"Error previewing file: {str(e)}")
            
            # Create download button
            zip_buffer = create_download_zip(temp_dir)
            
            st.download_button(
                label="📥 Download Results (ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"lysna_analysis_results_{language_choice.lower()}.zip",
                mime="application/zip",
                type="secondary",
                use_container_width=True
            )
            
            # Add button to clear results and start over
            if st.button("🔄 Start New Analysis", type="secondary"):
                # Clean up temporary directory
                import shutil
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
                
                # Clear session state
                for key in list(st.session_state.keys()):
                    if isinstance(key, str) and key.startswith(('analysis_complete', 'temp_dir', 'success_message', 'show_preview_', 'language_choice')):
                        del st.session_state[key]
                st.rerun()
    
    # Footer
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #666;'>
            <p>Lysna Assessment Analysis Tool </p>
        </div>
        """, 
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
