# Import necessary libraries
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import re
import textwrap
import matplotlib.ticker as ticker
from matplotlib.ticker import MaxNLocator

import lysna
import lysna.language

def sanitize_filename(name):
    return re.sub(r'[\/:*?"<>|]', '_', name)

# Function to wrap text
def wrap_text(text, width=65):
    return "\n".join(textwrap.wrap(text, width))

def bar_plots(data, color, output_dir, team, dynamic: bool = False) -> None:
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
        plt.ylabel('Count')
        wrapped_title = wrap_text(item)
        plt.title(wrapped_title)
        plt.gca().spines['top'].set_visible(False)
        plt.gca().spines['right'].set_visible(False)
        plt.gca().yaxis.set_major_locator(MaxNLocator(integer=True))
        sanitized_item = sanitize_filename(item)
        plt.savefig(f'{output_dir}/{sanitized_item}_{team}.png', transparent=True)
        plt.close()
        

def scale_plots(data: pd.DataFrame, scale: str, team: str, dir: str, color: str = "gray", dynamic: bool = False, lang = None) -> None:
    plt.figure(figsize=(10, 6))
    values = data.iloc[:-1][scale]
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
        
    if lang:
        xlabel = lang.scales["org_core"] if scale == "org_core" else lang.scales["team_core"] if scale == "team_core" else lang.scales["safety"] if scale == "safety" else lang.scales["dependability"]
    else:
        xlabel = "Organisational Core" if scale == "org_core" else "Team Core" if scale == "team_core" else "Psychological Safety" if scale == "safety" else "Professional Dependability"
    plt.xlabel(xlabel)
    plt.xticks(np.arange(1, 6, 1))
    plt.ylabel('Count')
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)
    plt.gca().yaxis.set_major_locator(MaxNLocator(integer=True))
    plt.savefig(f'{dir}/{scale}_{team}.png', transparent=True)
    plt.close()

def ranked_bar_plot(data: pd.DataFrame, scale: str, team: str, dir: str, color: str = "gray", dynamic: bool = False) -> None:
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

def get_bar_colors(values, default_color="gray", dynamic: bool = True) -> list:
    """Helper function to get colors based on values"""
    if not dynamic:
        return default_color
    return ['#ff9999' if x < 3 else '#90EE90' if x > 4 else default_color for x in values]
